# ===== ANALISADOR SPED EFD CONTRIBUIÇÕES v2.3 (BLOCO C COM C010 E MAIN) =====

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
    line = line.strip()
    if line and line.startswith('|') and line.endswith('|'):
        return line[1:-1].split('|')
    return []

def safe_get(data_list, index, default=''):
    return data_list[index] if len(data_list) > index else default

def safe_float_convert(value_str):
    if not value_str: return 0.0
    try:
        return float(value_str.replace('.', '').replace(',', '.'))
    except (ValueError, TypeError):
        return 0.0

def format_date(date_str):
    if date_str and len(date_str) == 8 and date_str.isdigit():
        try:
            return datetime.strptime(date_str, '%d%m%Y').strftime('%d/%m/%Y')
        except ValueError:
            return ''
    return ''

# Dicionários para traduzir códigos em textos legíveis
TIPO_ITEM_MAP = {
    '00': '00 - Mercadoria para Revenda', '01': '01 - Matéria-Prima',
    '02': '02 - Embalagem', '03': '03 - Produto em Processo',
    '04': '04 - Produto Acabado', '05': '05 - Subproduto',
    '06': '06 - Produto Intermediário', '07': '07 - Material de Uso e Consumo',
    '08': '08 - Ativo Imobilizado', '09': '09 - Serviços',
    '10': '10 - Outros Insumos', '99': '99 - Outras'
}

def processar_bloco_nota_contrib(c100_data, temp_c170_list, temp_c190_list, mapas):
    """
    Função auxiliar que processa os itens de UMA nota fiscal, aplicando a lógica condicional.
    """
    linhas_geradas = []
    mapa_produtos = mapas.get('produtos', {})

    if temp_c170_list:
        for campos in temp_c170_list:
            cod_item = safe_get(campos, 2)
            produto = mapa_produtos.get(cod_item, {})
            tipo_item_cod = produto.get('Tipo Item', '')
            
            row = {
                **c100_data,
                'Origem': 'C170 (Item)',
                'Código Item': cod_item,
                'Descrição Item': produto.get('Descrição Item', ''),
                'Código Barra': produto.get('Código Barra', ''),
                'Tipo Item': TIPO_ITEM_MAP.get(tipo_item_cod, tipo_item_cod),
                'NCM': produto.get('NCM', ''),
                'CFOP': safe_get(campos, 10),
                'CST ICMS': safe_get(campos, 9),
                'Qtde Item': safe_float_convert(safe_get(campos, 4)),
                'Vlr Item': safe_float_convert(safe_get(campos, 6)),
                'CST PIS': safe_get(campos, 24),
                'Vlr Base Cálculo PIS': safe_float_convert(safe_get(campos, 25)),
                'Alíquota PIS (%)': safe_float_convert(safe_get(campos, 26)),
                'Vlr PIS': safe_float_convert(safe_get(campos, 29)),
                'CST COFINS': safe_get(campos, 30),
                'Vlr Base Cálculo COFINS': safe_float_convert(safe_get(campos, 31)),
                'Alíquota COFINS (%)': safe_float_convert(safe_get(campos, 32)),
                'Vlr COFINS': safe_float_convert(safe_get(campos, 35)),
                'Conta Contábil': safe_get(campos, 36),
            }
            linhas_geradas.append(row)
    elif temp_c190_list:
        for campos in temp_c190_list:
            cfop = safe_get(campos, 2)
            row = {
                **c100_data,
                'Origem': 'C190 (Resumo)',
                'CST ICMS': safe_get(campos, 1),
                'CFOP': cfop,
                'Vlr Operação': safe_float_convert(safe_get(campos, 4)),
            }
            linhas_geradas.append(row)
    
    return linhas_geradas

def gerar_relatorio_efd_contrib(lista_arquivos):
    """Função principal adaptada para usar o CNPJ do registro C010."""
    
    mapa_empresa, mapa_participantes, mapa_produtos = {}, {}, {}
    
    print("Iniciando Etapa 1: Mapeamento de dados de apoio...")
    for filepath in lista_arquivos:
        nome_arquivo = os.path.basename(filepath)
        print(f"  - Mapeando arquivo: {nome_arquivo}")
        with open(filepath, 'r', encoding='latin-1') as f:
            for line in f:
                campos = parse_line(line)
                if not campos: continue
                reg = campos[0]
                if reg == '0000' and nome_arquivo not in mapa_empresa:
                    mapa_empresa[nome_arquivo] = {'Nome Empresa Declarante': safe_get(campos, 6), 'Período': f"{format_date(safe_get(campos, 4))} a {format_date(safe_get(campos, 5))}"}
                elif reg == '0150':
                    cod_part = safe_get(campos, 1)
                    if cod_part and cod_part not in mapa_participantes:
                        mapa_participantes[cod_part] = {'Nome Participante': safe_get(campos, 2)}
                elif reg == '0200':
                    cod_item = safe_get(campos, 1)
                    if cod_item and cod_item not in mapa_produtos:
                        mapa_produtos[cod_item] = {
                            'Descrição Item': safe_get(campos, 2), 
                            'Código Barra': safe_get(campos, 3),
                            'Tipo Item': safe_get(campos, 6), 
                            'NCM': safe_get(campos, 7)
                        }
    
    mapas_gerais = {'produtos': mapa_produtos}
    report_rows = []
    print("\nIniciando Etapa 2: Processamento do Bloco C...")
    for filepath in lista_arquivos:
        nome_arquivo = os.path.basename(filepath)
        print(f"  - Processando arquivo: {nome_arquivo}")
        info_empresa = mapa_empresa.get(nome_arquivo, {})
        
        with open(filepath, 'r', encoding='latin-1') as f:
            current_c100_data = {}
            temp_c170_list = []
            temp_c190_list = []
            cnpj_estabelecimento_atual = ""
            
            for line in f:
                campos = parse_line(line)
                if not campos: continue
                reg = campos[0]

                if reg == 'C010':
                    cnpj_estabelecimento_atual = safe_get(campos, 1)
                
                if (reg == 'C100' or not reg.startswith('C1')) and current_c100_data:
                    report_rows.extend(processar_bloco_nota_contrib(current_c100_data, temp_c170_list, temp_c190_list, mapas_gerais))
                    current_c100_data, temp_c170_list, temp_c190_list = {}, [], []

                if reg == 'C100':
                    cod_part = safe_get(campos, 3)
                    participante = mapa_participantes.get(cod_part, {})
                    current_c100_data = {
                        'CNPJ Estabelecimento': cnpj_estabelecimento_atual,
                        'Nome Empresa Declarante': info_empresa.get('Nome Empresa Declarante', ''),
                        'Período': info_empresa.get('Período', ''),
                        'Tipo Operação': 'Entrada' if safe_get(campos, 1) == '0' else 'Saída',
                        'Código Participante': cod_part,
                        'Nome Participante': participante.get('Nome Participante', ''),
                        'Número Documento': safe_get(campos, 7),
                        'Chave NF-e': safe_get(campos, 8),
                        'Data Documento': format_date(safe_get(campos, 10)),
                        'Vlr Documento': safe_float_convert(safe_get(campos, 12)),
                    }
                
                elif reg == 'C170' and current_c100_data:
                    temp_c170_list.append(campos)
                
                elif reg == 'C190' and current_c100_data:
                    temp_c190_list.append(campos)
            
            if current_c100_data:
                report_rows.extend(processar_bloco_nota_contrib(current_c100_data, temp_c170_list, temp_c190_list, mapas_gerais))

    if not report_rows:
        raise ValueError("Nenhum registro C170 ou C190 foi encontrado nos arquivos.")

    df_final = pd.DataFrame(report_rows)
    
    colunas_finais = ['CNPJ Estabelecimento', 'Nome Empresa Declarante', 'Período', 'Data Documento', 'Tipo Operação', 'Número Documento', 'Chave NF-e', 'Código Participante', 'Nome Participante', 'Vlr Documento', 'Origem', 'Código Item', 'Descrição Item', 'Código Barra', 'Tipo Item', 'NCM', 'CFOP', 'CST ICMS', 'Qtde Item', 'Vlr Item', 'CST PIS', 'Vlr Base Cálculo PIS', 'Alíquota PIS (%)', 'Vlr PIS', 'CST COFINS', 'Vlr Base Cálculo COFINS', 'Alíquota COFINS (%)', 'Vlr COFINS', 'Conta Contábil']
    for col in colunas_finais:
        if col not in df_final.columns:
            df_final[col] = ''
            
    return df_final[colunas_finais]

# ==============================================================================
# PARTE 2: INTERFACE GRÁFICA (PyQt6)
# ==============================================================================

class SpedContribApp(QWidget):
    def __init__(self):
        super().__init__()
        self.selected_files = []
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Analisador EFD Contribuições - Bloco C (com C010)")
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
        self.run_button = QPushButton("🚀 Gerar Relatório do Bloco C")
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
        self.status_label.setText("Processando múltiplos arquivos... Isso pode levar um momento.")
        QApplication.processEvents()
        try:
            df_result = gerar_relatorio_efd_contrib(self.selected_files)
            self.save_report(df_result)
            self.status_label.setText("✅ Relatório consolidado gerado com sucesso!")
        except Exception as e:
            self.status_label.setText(f"❌ Erro: {e}")
            QMessageBox.critical(self, "Erro de Processamento", f"Ocorreu um erro.\n\nDetalhes: {e}")
        finally:
            self.run_button.setEnabled(True)

    def save_report(self, dataframe):
        default_filename = f"Relatorio_EFD_Contrib_Bloco_C_{datetime.now().strftime('%Y%m%d')}.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(self, "Salvar Relatório", default_filename, "Arquivos Excel (*.xlsx)")
        if save_path:
            try:
                self.status_label.setText("Salvando arquivo Excel...")
                QApplication.processEvents()
                dataframe.to_excel(save_path, index=False, sheet_name="Bloco_C_Consolidado")
                QMessageBox.information(self, "Sucesso", f"Relatório salvo em:\n{save_path}")
            except Exception as e:
                self.status_label.setText(f"❌ Erro ao salvar o arquivo: {e}")
                QMessageBox.critical(self, "Erro ao Salvar", f"Não foi possível salvar o arquivo.\n\nDetalhes: {e}")

# ==============================================================================
# PARTE 3: PONTO DE ENTRADA
# ==============================================================================

def main():
    """Função principal que inicia esta aplicação específica."""
    # Reutiliza a instância da aplicação principal (do Hub) se ela existir.
    # Caso contrário, cria uma nova para execução independente.
    app = QApplication.instance() or QApplication(sys.argv)
    
    # Cria e exibe a janela da aplicação.
    # **ALTERADO: Nome da classe para SpedContribApp**
    window = SpedContribApp()
    window.show()

    # Retorna a aplicação e a janela para o chamador poder gerenciá-las.
    return app, window

# Este bloco permite que o script ainda seja executável de forma independente para testes
if __name__ == '__main__':
    app, window = main()
    sys.exit(app.exec())