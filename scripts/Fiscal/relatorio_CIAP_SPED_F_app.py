# ===== ANALISADOR SPED FISCAL v2.3 - BLOCO G (CIAP COM DESCRIÇÃO DO 0300) (MODIFICADO PARA O HUB) =====

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

# Dicionário para traduzir o tipo de movimentação do CIAP
TIPO_MOV_MAP = {
    'MC': 'MC - Imobilização de Mercadoria para Uso/Consumo',
    'IM': 'IM - Imobilização de Insumo/Matéria-Prima',
    'CI': 'CI - Conclusão de Obra (componente)',
    'SI': 'SI - Saldo Inicial de Bens',
    'AT': 'AT - Imobilização vinda de Ativo Circulante',
    'PE': 'PE - Imobilização de bem objeto de Contrato de Arrendamento',
    'OU': 'OU - Outras Entradas'
}

def gerar_relatorio_bloco_g(lista_arquivos):
    """Função principal que implementa a lógica de extração para o Bloco G."""
    
    # --- ETAPA 1: Mapear dados de apoio de todos os arquivos ---
    mapa_empresa, mapa_participantes, mapa_bens_ciap = {}, {}, {}
    
    print("Iniciando Etapa 1: Mapeamento de dados de apoio (0000, 0150, 0300)...")
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
                        mapa_participantes[cod_part] = {'Nome Participante': safe_get(campos, 2)}
                elif reg == '0300':
                    cod_bem = safe_get(campos, 1)
                    if cod_bem and cod_bem not in mapa_bens_ciap:
                        mapa_bens_ciap[cod_bem] = {'Descrição do Bem/Componente': safe_get(campos, 3)}
    
    # --- ETAPA 2: Processar Bloco G e montar o relatório ---
    report_rows = []
    print("\nIniciando Etapa 2: Processamento do Bloco G...")
    for filepath in lista_arquivos:
        nome_arquivo = os.path.basename(filepath)
        print(f"  - Processando arquivo: {nome_arquivo}")
        info_empresa = mapa_empresa.get(nome_arquivo, {})
        
        with open(filepath, 'r', encoding='latin-1') as f:
            current_g125_data = {}
            
            for line in f:
                campos = parse_line(line)
                if not campos: continue
                reg = campos[0]
                
                if not reg.startswith('G13') and not reg.startswith('G14') and current_g125_data:
                    report_rows.append(current_g125_data)
                    current_g125_data = {}

                if reg == 'G125':
                    cod_item = safe_get(campos, 1)
                    bem_info = mapa_bens_ciap.get(cod_item, {})
                    tipo_mov_cod = safe_get(campos, 3)
                    
                    current_g125_data = {
                        **info_empresa,
                        'Código do Bem/Componente': cod_item,
                        'Descrição do Bem/Componente': bem_info.get('Descrição do Bem/Componente', 'BEM NÃO CADASTRADO NO REG 0300'),
                        'Data da Movimentação': format_date(safe_get(campos, 2)),
                        'Tipo de Movimentação': TIPO_MOV_MAP.get(tipo_mov_cod, tipo_mov_cod),
                        'Valor ICMS da Operação': safe_float_convert(safe_get(campos, 4)),
                        'Valor Parcela Apropriável': safe_float_convert(safe_get(campos, 9)),
                        'Nº Parcela / Total Parcelas': safe_get(campos, 8),
                    }
                elif reg == 'G130' and current_g125_data:
                    cod_part = safe_get(campos, 2)
                    participante = mapa_participantes.get(cod_part, {})
                    
                    current_g125_data.update({
                        'Nome do Participante': participante.get('Nome Participante', ''),
                        'Modelo do Documento': safe_get(campos, 3),
                        'Série': safe_get(campos, 4),
                        'Número do Documento': safe_get(campos, 5),
                        'Chave da NF-e': safe_get(campos, 6),
                    })
                    report_rows.append(current_g125_data)
                    current_g125_data = {}

            if current_g125_data:
                report_rows.append(current_g125_data)

    if not report_rows:
        raise ValueError("Nenhum registro de movimentação de CIAP (G125) foi encontrado nos arquivos.")

    df_final = pd.DataFrame(report_rows)
    
    colunas_finais = [
        'CNPJ', 'Inscrição Estadual', 'Período', 'Data da Movimentação', 
        'Código do Bem/Componente', 'Descrição do Bem/Componente', 'Tipo de Movimentação', 
        'Nome do Participante', 'Modelo do Documento', 'Série', 'Número do Documento', 'Chave da NF-e',
        'Valor ICMS da Operação', 'Nº Parcela / Total Parcelas', 'Valor Parcela Apropriável'
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

class SpedCiapApp(QWidget):
    def __init__(self):
        super().__init__()
        self.selected_files = []
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Gerador de Relatório SPED Fiscal - Bloco G (CIAP)")
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
        
        self.run_button = QPushButton("🚀 Gerar Relatório CIAP (Bloco G)")
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
            df_result = gerar_relatorio_bloco_g(self.selected_files)
            self.save_report(df_result)
            self.status_label.setText("✅ Relatório consolidado do Bloco G gerado com sucesso!")
        except Exception as e:
            self.status_label.setText(f"❌ Erro: {e}")
            QMessageBox.critical(self, "Erro de Processamento", f"Ocorreu um erro.\n\nDetalhes: {e}")
        finally:
            self.run_button.setEnabled(True)

    def save_report(self, dataframe):
        default_filename = f"Relatorio_Bloco_G_CIAP_{datetime.now().strftime('%Y%m%d')}.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(self, "Salvar Relatório", default_filename, "Arquivos Excel (*.xlsx)")
        
        if save_path:
            try:
                self.status_label.setText("Salvando arquivo Excel...")
                QApplication.processEvents()
                dataframe.to_excel(save_path, index=False, sheet_name="Bloco_G_CIAP")
                QMessageBox.information(self, "Sucesso", f"Relatório salvo em:\n{save_path}")
            except Exception as e:
                self.status_label.setText(f"❌ Erro ao salvar o arquivo: {e}")
                QMessageBox.critical(self, "Erro ao Salvar", f"Não foi possível salvar o arquivo.\n\nDetalhes: {e}")

# ==============================================================================
# PARTE 3: PONTO DE ENTRADA (MODIFICADO)
# ==============================================================================

def main():
    """Função principal que inicia esta aplicação específica."""
    # Reutiliza a instância da aplicação principal (do Hub) se ela existir.
    # Caso contrário, cria uma nova para execução independente.
    app = QApplication.instance() or QApplication(sys.argv)

    # Cria e exibe a janela da aplicação.
    window = SpedCiapApp()
    window.show()

    # Retorna a aplicação e a janela para o chamador poder gerenciá-las.
    return app, window

# Este bloco permite que o script ainda seja executável de forma independente para testes
if __name__ == '__main__':
    app, window = main()
    sys.exit(app.exec())