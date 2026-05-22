# Caminho do Código: /scripts/Analise/analise_fiscal_app.py
# (MODIFICADO V8) Remoção de emojis dos botões para um visual mais profissional.
# Mantido apenas o símbolo '⋮' para o menu de opções.

import sys
import os
import pandas as pd
import numpy as np
import traceback # Importado para imprimir o erro completo

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QTreeWidget, QTreeWidgetItem, QSplitter, QLabel,
    QPushButton, QFileDialog, QMessageBox, QHeaderView,
    QListWidget, QStackedWidget, QMenu, QTabWidget
)
from PyQt6.QtCore import Qt, pyqtSignal, QSize
from PyQt6.QtGui import QFont, QColor, QAction

# --- Integração do Matplotlib com PyQt6 ---
import matplotlib
matplotlib.use('Qt5Agg')
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.figure import Figure
import matplotlib.pyplot as plt

# --- (CORREÇÃO DE IMPORTAÇÃO) ---
try:
    # Tenta importar relativamente (funciona quando chamado pelo hub)
    from . import analise_fiscal_logic as logic
except ImportError:
    try:
        # Tenta importar como se estivesse na raiz do projeto (funciona no teste local)
        from scripts.Analise import analise_fiscal_logic as logic
    except ImportError:
        try:
            # Tenta adicionar o diretório pai ao path (outra tentativa local)
            current_dir = os.path.dirname(os.path.abspath(__file__))
            parent_dir = os.path.dirname(current_dir)
            root_dir = os.path.dirname(parent_dir) # HUB-Auditoria directory
            if root_dir not in sys.path:
                sys.path.insert(0, root_dir)
            from scripts.Analise import analise_fiscal_logic as logic
        except ImportError as e:
            print(f"Erro fatal: Não foi possível importar 'analise_fiscal_logic'. {e}")
            app = QApplication.instance() or QApplication(sys.argv)
            QMessageBox.critical(None, "Erro Fatal de Importação",
                "Não foi possível encontrar o arquivo 'analise_fiscal_logic.py'.\n\n"
                f"Verifique se o arquivo está em 'scripts/Analise/'.\n\nErro: {e}")
            sys.exit(1)
# --- (FIM DA CORREÇÃO) ---


# Define cores padrão (estilo Google)
STYLE_COLORS = {
    "background": "#FFFFFF",
    "text": "#333333",
    "primary_accent": "#4285F4", # Azul
    "secondary_accent": "#34A853", # Verde
    "warning_accent": "#DB4437", # Vermelho
    "header_bg": "#F1F3F4", # Cinza claro
    "grid": "#E0E0E0",
    "border": "#DADCE0", # Cinza borda
    "charts": ['#4285F4', '#DB4437', '#F4B400', '#0F9D58', '#AB47BC', '#00ACC1', '#FF7043', '#7E57C2']
}

# Stylesheet para um visual clean
STYLESHEET = f"""
    QWidget {{
        font-family: 'Arial', 'Segoe UI';
        font-size: 10pt;
        background-color: {STYLE_COLORS['background']};
        color: {STYLE_COLORS['text']};
    }}
    QTreeWidget QHeaderView::section {{
        font-weight: bold;
        background-color: {STYLE_COLORS['header_bg']};
        padding: 6px;
        border: 1px solid {STYLE_COLORS['grid']};
    }}
    QTreeWidget {{
        border: 1px solid {STYLE_COLORS['grid']};
        alternate-background-color: #FDFDFD;
    }}
    QTreeWidget::item {{
        padding: 3px 0px;
    }}
    QTreeWidget::item:selected {{
        background-color: {STYLE_COLORS['primary_accent']};
        color: white;
    }}
    QLabel#TitleLabel {{
        font-size: 13pt;
        font-weight: bold;
    }}

    /* --- Estilos de Botão Padronizados --- */
    QPushButton {{
        border-radius: 4px;
        font-weight: 500; /* Levemente mais pesado que normal */
    }}

    /* Botão Primário (Azul) */
    QPushButton#PrimaryButton {{
        background-color: {STYLE_COLORS['primary_accent']};
        color: white;
        border: none;
        padding: 8px 12px;
    }}
    QPushButton#PrimaryButton:hover {{
        background-color: #3367D6;
    }}

    /* Botão Secundário (Verde) */
    QPushButton#SecondaryButton {{
        background-color: {STYLE_COLORS['secondary_accent']};
        color: white;
        border: none;
        padding: 8px 12px;
    }}
    QPushButton#SecondaryButton:hover {{
        background-color: #2B8B45;
    }}

    /* Botão de Utilidade (Cinza) */
    QPushButton#UtilityButton {{
        background-color: {STYLE_COLORS['header_bg']};
        color: {STYLE_COLORS['text']};
        border: 1px solid {STYLE_COLORS['border']};
        padding: 8px 16px;
    }}
    QPushButton#UtilityButton:hover {{
        background-color: #E8EAED;
        border-color: #C6C6C6;
    }}

    /* Botão de Menu '⋮' (específico) */
    QPushButton#OptionsMenuButton {{
        font-size: 14pt;
        font-weight: bold;
        padding: 0px 8px; /* Padding horizontal apenas */
        min-height: 38px; /* Altura fixa para alinhar com outros botões do header */
        max-width: 45px;
        background-color: {STYLE_COLORS['header_bg']};
        color: {STYLE_COLORS['text']};
        border: 1px solid {STYLE_COLORS['border']};
        border-radius: 4px;
    }}
     QPushButton#OptionsMenuButton:hover {{
        background-color: #E8EAED;
        border-color: #C6C6C6;
    }}
    /* --- Fim dos Estilos de Botão --- */

    QListWidget {{
        border: 1px solid {STYLE_COLORS['border']};
        border-radius: 4px;
    }}

    /* Abas dos Gráficos */
    QTabWidget::pane {{
        border: 1px solid {STYLE_COLORS['grid']};
        border-top: 0px;
    }}
    QTabBar::tab {{
        background: {STYLE_COLORS['header_bg']};
        border: 1px solid {STYLE_COLORS['grid']};
        border-bottom: none;
        padding: 10px 18px;
        font-weight: bold;
    }}
    QTabBar::tab:selected {{
        background: {STYLE_COLORS['background']};
        border-bottom: 2px solid {STYLE_COLORS['primary_accent']};
    }}
"""

# ==============================================================================
# CLASSE AUXILIAR PARA ORDENAÇÃO NUMÉRICA DA TABELA
# ==============================================================================

class NumericTreeWidgetItem(QTreeWidgetItem):
    """
    Subclasse de QTreeWidgetItem para permitir a ordenação numérica correta
    da coluna 'Valor Total'.
    """
    def __lt__(self, other):
        column = self.treeWidget().sortColumn()
        if column == 4: # Coluna 'Valor Total'
            self_value = self.data(column, Qt.ItemDataRole.UserRole)
            other_value = other.data(column, Qt.ItemDataRole.UserRole)
            self_value = self_value if self_value is not None else -float('inf')
            other_value = other_value if other_value is not None else -float('inf')
            return self_value < other_value
        else:
            return super().__lt__(other)

# ==============================================================================
# PÁGINA 1: WIDGET DO MENU DE SELEÇÃO DE SPED
# ==============================================================================

class MenuWidget(QWidget):
    """
    Esta é a primeira tela que o usuário vê.
    Oferece duas opções: carregar TXT ou carregar CSV/XLSX.
    """
    data_processed = pyqtSignal(pd.DataFrame)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.selected_files = []
        self.initUI()

    def initUI(self):
        # --- Layout com mais espaçamento ---
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(30, 30, 30, 30)

        main_layout.addWidget(QLabel("1. Selecione um ou mais arquivos EFD Contribuições (.txt):"))
        self.file_list_widget = QListWidget()
        self.file_list_widget.setMinimumHeight(120)
        main_layout.addWidget(self.file_list_widget)

        # (MODIFICADO V8) Removido emoji
        browse_button = QPushButton("Procurar Arquivos...")
        browse_button.setObjectName("UtilityButton")
        browse_button.setCursor(Qt.CursorShape.PointingHandCursor)
        browse_button.clicked.connect(self.select_files)

        # (MODIFICADO V8) Removido emoji
        self.run_button = QPushButton("Gerar Análise Consolidada do Bloco C (do TXT)")
        self.run_button.setFont(QFont("Segoe UI", 11))
        self.run_button.setObjectName("SecondaryButton") # Verde
        self.run_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.run_button.clicked.connect(self.start_processing_txt)

        # (MODIFICADO V8) Removido emoji
        self.data_file_button = QPushButton("Carregar Relatório (CSV ou XLSX)")
        self.data_file_button.setObjectName("PrimaryButton") # Azul
        self.data_file_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.data_file_button.clicked.connect(self.load_data_file)

        # Layout para botão procurar
        browse_layout = QHBoxLayout()
        browse_layout.addWidget(browse_button)
        browse_layout.addStretch()
        main_layout.addLayout(browse_layout)

        main_layout.addSpacing(15)
        main_layout.addWidget(self.run_button)
        main_layout.addWidget(QLabel("ou..."), alignment=Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(self.data_file_button)
        main_layout.addSpacing(10)

        self.status_label = QLabel("Pronto para começar.")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(self.status_label)

        main_layout.addStretch()
        # --- Fim Layout ---

    def select_files(self):
        filepaths, _ = QFileDialog.getOpenFileNames(self, "Selecione os arquivos EFD Contribuições", "", "Arquivos de Texto (*.txt)")
        if filepaths:
            self.selected_files = filepaths
            self.file_list_widget.clear()
            for file_path in self.selected_files:
                self.file_list_widget.addItem(os.path.basename(file_path))

    def start_processing_txt(self):
        """
        Processa os arquivos TXT selecionados.
        """
        if not self.selected_files:
            QMessageBox.warning(self, "Atenção", "Nenhum arquivo TXT foi selecionado.")
            return

        self.run_button.setEnabled(False)
        self.status_label.setText("Processando múltiplos arquivos SPED... Isso pode levar um momento.")
        QApplication.processEvents()

        try:
            df_result = logic.gerar_relatorio_efd_contrib(self.selected_files)
            self.status_label.setText("Dados processados! Carregando dashboard...") # Removido emoji
            QApplication.processEvents()
            self.data_processed.emit(df_result)

        except Exception as e:
            self.status_label.setText(f"Erro: {e}") # Removido emoji
            QMessageBox.critical(self, "Erro de Processamento", f"Ocorreu um erro ao processar os arquivos TXT.\n\nDetalhes: {e}")
        finally:
            self.run_button.setEnabled(True)

    def load_data_file(self):
        """
        Carrega um arquivo de dados pré-processado (.csv ou .xlsx).
        """
        filepaths, _ = QFileDialog.getOpenFileNames(
            self,
            "Selecione o Relatório (CSV ou XLSX)",
            "",
            "Arquivos de Dados (*.csv *.xlsx)"
        )
        if not filepaths:
            return

        self.data_file_button.setEnabled(False)
        self.status_label.setText(f"Carregando arquivo: {os.path.basename(filepaths[0])}...")
        QApplication.processEvents()

        try:
            df_raw = pd.DataFrame()
            single_filepath = filepaths[0]

            if single_filepath.endswith('.csv'):
                df_raw = pd.read_csv(single_filepath, sep=',', encoding='latin-1', low_memory=False)
            elif single_filepath.endswith('.xlsx'):
                try:
                    df_raw = pd.read_excel(single_filepath, sheet_name="Bloco_C_Consolidado")
                except Exception:
                    df_raw = pd.read_excel(single_filepath, sheet_name=0)
            else:
                raise ValueError("Formato de arquivo não suportado. Selecione .csv ou .xlsx")

            self.status_label.setText("Arquivo carregado! Carregando dashboard...") # Removido emoji
            QApplication.processEvents()
            self.data_processed.emit(df_raw)

        except Exception as e:
            self.status_label.setText(f"Erro: {e}") # Removido emoji
            QMessageBox.critical(self, "Erro ao Carregar Arquivo", f"Não foi possível ler o arquivo.\n\nDetalhes: {e}")
        finally:
            self.data_file_button.setEnabled(True)

# ==============================================================================
# PÁGINA 2: WIDGET DO DASHBOARD
# ==============================================================================

class DashboardWidget(QWidget):
    """
    Esta é a segunda tela (o dashboard).
    """
    back_to_menu_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)

        self.df_base = pd.DataFrame()
        self.df_suppliers_abc = pd.DataFrame()

        self.selected_supplier_cnpj = None
        self.selected_ncm = None
        self.selected_cfop = None

        self.is_populating = False
        self.ncm_donut_cid = None
        self.cfop_donut_cid = None

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(15, 15, 15, 15)

        # --- Header com botões 'Utility' ---
        header_layout = QHBoxLayout()
        header_layout.setSpacing(10)

        # Botão de Opções "⋮"
        self.btn_options = QPushButton("⋮")
        self.btn_options.setObjectName("OptionsMenuButton") # Estilo específico
        self.btn_options.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_options.clicked.connect(self.show_export_menu)
        header_layout.addWidget(self.btn_options)

        title_label = QLabel("Dashboard de Análise Fiscal - Entradas (Fornecedores x NCM x CFOP)")
        title_label.setObjectName("TitleLabel")
        header_layout.addWidget(title_label)
        header_layout.addStretch()

        # (MODIFICADO V8) Removido emoji
        self.btn_back = QPushButton("Voltar")
        self.btn_back.setObjectName("UtilityButton")
        self.btn_back.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_back.clicked.connect(self.request_back_to_menu)
        header_layout.addWidget(self.btn_back)

        # (MODIFICADO V8) Removido emoji
        self.btn_reset = QPushButton("Limpar Filtros")
        self.btn_reset.setObjectName("UtilityButton")
        self.btn_reset.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_reset.clicked.connect(self.reset_all_filters)
        header_layout.addWidget(self.btn_reset)

        main_layout.addLayout(header_layout)
        # --- Fim do Header ---

        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(self.splitter, 1)

        # --- Painel da Esquerda (Tabela de Fornecedores) ---
        supplier_frame = QWidget()
        supplier_layout = QVBoxLayout(supplier_frame)
        supplier_layout.setContentsMargins(0, 5, 5, 0)
        supplier_layout.setSpacing(10)

        lbl_supplier = QLabel("Fornecedores (Curva ABC)")
        lbl_supplier.setObjectName("TitleLabel")
        supplier_layout.addWidget(lbl_supplier)

        self.create_supplier_table()
        supplier_layout.addWidget(self.tree_suppliers)
        self.splitter.addWidget(supplier_frame)
        # --- Fim Painel da Esquerda ---

        # --- Painel da Direita (Gráficos) ---
        charts_frame = QWidget()
        charts_layout = QVBoxLayout(charts_frame)
        charts_layout.setContentsMargins(5, 5, 0, 0)
        charts_layout.setSpacing(10)

        charts_splitter = QSplitter(Qt.Orientation.Vertical)

        # Abas para os gráficos de Rosca (NCM e CFOP)
        self.charts_tab_widget = QTabWidget()
        self.init_ncm_donut_chart()
        self.init_cfop_donut_chart()

        self.charts_tab_widget.addTab(self.ncm_donut_frame, "Análise por NCM")
        self.charts_tab_widget.addTab(self.cfop_donut_frame, "Análise por CFOP")

        charts_splitter.addWidget(self.charts_tab_widget)

        # Gráfico de Linha (Tempo)
        self.init_line_chart()
        charts_splitter.addWidget(self.line_frame)

        charts_splitter.setSizes([int(self.height() * 0.6), int(self.height() * 0.4)])
        charts_layout.addWidget(charts_splitter)

        self.splitter.addWidget(charts_frame)
        # --- Fim Painel da Direita ---

        self.splitter.setSizes([int(self.width() * 0.45), int(self.width() * 0.55)])

    # --- Funções para o Menu de Exportação ---

    def show_export_menu(self):
        """Cria e exibe o menu de exportação no botão '...'."""
        menu = QMenu(self)

        # (MODIFICADO V8) Removido emoji
        export_excel_action = QAction("Exportar Análise para Excel...", self)
        export_excel_action.triggered.connect(self.export_analysis_to_excel)
        menu.addAction(export_excel_action)

        # (MODIFICADO V8) Removido emoji
        export_pdf_action = QAction("Exportar Gráficos para PDF...", self)
        export_pdf_action.triggered.connect(self.export_charts_to_pdf)
        menu.addAction(export_pdf_action)

        button_point = self.btn_options.mapToGlobal(self.btn_options.rect().bottomLeft())
        menu.exec(button_point)

    def get_currently_filtered_data(self):
        """Retorna um DataFrame filtrado com base nos seletores atuais."""
        df_current = self.df_base.copy()
        if self.selected_supplier_cnpj:
            df_current = df_current[df_current['CNPJ/CPF Participante'] == self.selected_supplier_cnpj]
        if self.selected_ncm:
            df_current = df_current[df_current['NCM'] == self.selected_ncm]
        if self.selected_cfop:
            df_current = df_current[df_current['CFOP'] == self.selected_cfop]
        return df_current

    def export_analysis_to_excel(self):
        """Exporta os resumos da análise atual para um arquivo Excel."""
        if self.df_base.empty:
            QMessageBox.warning(self, "Sem Dados", "Não há dados para exportar.")
            return

        default_filename = "Analise_Fiscal_Dashboard.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(self, "Salvar Análise", default_filename, "Arquivos Excel (*.xlsx)")

        if not save_path:
            return

        try:
            df_filtered = self.get_currently_filtered_data()

            df_suppliers_export = logic.calculate_supplier_abc(df_filtered)
            df_ncm_export = logic.calculate_ncm_summary(df_filtered)
            df_cfop_export = logic.calculate_cfop_summary(df_filtered)
            df_monthly_export = logic.calculate_monthly_summary(df_filtered)

            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                df_suppliers_export.to_excel(writer, sheet_name="Curva ABC Fornecedores", index=False)
                df_ncm_export.to_excel(writer, sheet_name="Top NCM", index=False)
                df_cfop_export.to_excel(writer, sheet_name="Top CFOP", index=False)
                df_monthly_export.to_excel(writer, sheet_name="Total Mensal", index=False)

            QMessageBox.information(self, "Sucesso", f"Análise exportada com sucesso para:\n{save_path}")

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Exportar", f"Não foi possível salvar o arquivo Excel.\n\nDetalhe: {e}")
            traceback.print_exc()

    def export_charts_to_pdf(self):
        """Exporta os três gráficos (NCM, CFOP, Linha) para um único PDF."""
        if self.df_base.empty:
            QMessageBox.warning(self, "Sem Dados", "Não há dados para exportar.")
            return

        default_filename = "Relatorio_Graficos_Analise_Fiscal.pdf"
        save_path, _ = QFileDialog.getSaveFileName(self, "Salvar Gráficos em PDF", default_filename, "Arquivos PDF (*.pdf)")

        if not save_path:
            return

        try:
            with PdfPages(save_path) as pdf:
                # Página 1: Gráfico de NCM
                self.fig_ncm_donut.suptitle("Análise - Top NCMs", fontsize=16, y=1.05)
                pdf.savefig(self.fig_ncm_donut, bbox_inches='tight')
                self.fig_ncm_donut.suptitle(None)

                # Página 2: Gráfico de CFOP
                self.fig_cfop_donut.suptitle("Análise - Top CFOPs", fontsize=16, y=1.05)
                pdf.savefig(self.fig_cfop_donut, bbox_inches='tight')
                self.fig_cfop_donut.suptitle(None)

                # Página 3: Gráfico de Linha
                self.fig_line.suptitle("Análise - Compras por Período", fontsize=16, y=1.05)
                pdf.savefig(self.fig_line, bbox_inches='tight')
                self.fig_line.suptitle(None)

            QMessageBox.information(self, "Sucesso", f"Gráficos exportados com sucesso para:\n{save_path}")

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Exportar PDF", f"Não foi possível salvar o arquivo PDF.\n\nDetalhe: {e}")
            traceback.print_exc()

    # --- Fim das Funções de Exportação ---

    def request_back_to_menu(self):
        self.back_to_menu_requested.emit()

    def create_supplier_table(self):
        cols = ('Sel.', 'Curva', 'CNPJ/CPF Participante', 'Nome do Participante', 'Valor Total')
        self.tree_suppliers = QTreeWidget()
        self.tree_suppliers.setColumnCount(len(cols))
        self.tree_suppliers.setHeaderLabels(cols)
        self.tree_suppliers.setAlternatingRowColors(True)
        self.tree_suppliers.setSelectionMode(QTreeWidget.SelectionMode.SingleSelection)
        header = self.tree_suppliers.header()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents) # Sel.
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents) # Curva
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents) # CNPJ/CPF
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch) # Nome
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents) # Valor

        self.tree_suppliers.setSortingEnabled(True)
        self.tree_suppliers.sortByColumn(4, Qt.SortOrder.DescendingOrder)

        self.tree_suppliers.itemClicked.connect(self.on_supplier_select)

    def init_ncm_donut_chart(self):
        """Cria apenas o gráfico NCM."""
        self.ncm_donut_frame = QWidget()
        layout = QVBoxLayout(self.ncm_donut_frame)
        layout.setContentsMargins(5, 5, 5, 5)

        self.fig_ncm_donut = Figure(figsize=(5, 4), dpi=100)
        self.fig_ncm_donut.patch.set_facecolor(STYLE_COLORS['background'])
        self.ax_ncm_donut = self.fig_ncm_donut.add_subplot(111)
        self.ax_ncm_donut.set_facecolor(STYLE_COLORS['background'])
        self.canvas_ncm_donut = FigureCanvas(self.fig_ncm_donut)
        layout.addWidget(self.canvas_ncm_donut)
        self.canvas_ncm_donut.draw()
        self.ncm_donut_cid = self.fig_ncm_donut.canvas.mpl_connect('pick_event', self.on_ncm_click)

    def init_cfop_donut_chart(self):
        """Cria o widget para o gráfico de rosca de CFOP."""
        self.cfop_donut_frame = QWidget()
        layout = QVBoxLayout(self.cfop_donut_frame)
        layout.setContentsMargins(5, 5, 5, 5)

        self.fig_cfop_donut = Figure(figsize=(5, 4), dpi=100)
        self.fig_cfop_donut.patch.set_facecolor(STYLE_COLORS['background'])
        self.ax_cfop_donut = self.fig_cfop_donut.add_subplot(111)
        self.ax_cfop_donut.set_facecolor(STYLE_COLORS['background'])
        self.canvas_cfop_donut = FigureCanvas(self.fig_cfop_donut)
        layout.addWidget(self.canvas_cfop_donut)
        self.canvas_cfop_donut.draw()
        self.cfop_donut_cid = self.fig_cfop_donut.canvas.mpl_connect('pick_event', self.on_cfop_click)

    def init_line_chart(self):
        self.line_frame = QWidget()
        layout = QVBoxLayout(self.line_frame)
        layout.setContentsMargins(5, 15, 5, 5)
        lbl_line = QLabel("Compras por Período")
        lbl_line.setObjectName("TitleLabel")
        layout.addWidget(lbl_line)
        self.fig_line = Figure(figsize=(5, 3), dpi=100)
        self.fig_line.patch.set_facecolor(STYLE_COLORS['background'])
        self.ax_line = self.fig_line.add_subplot(111)
        self.ax_line.set_facecolor(STYLE_COLORS['background'])
        self.canvas_line = FigureCanvas(self.fig_line)
        layout.addWidget(self.canvas_line)
        self.canvas_line.draw()

    def populate_dashboard(self, df_raw):
        """
        Função central para popular o dashboard.
        """
        self.is_populating = True
        try:
            self.df_base = logic.process_dataframe_for_dashboard(df_raw)
            if self.df_base.empty:
                QMessageBox.information(self, "Sem Dados para Análise",
                    "Processamento concluído.\n\n"
                    "Não foram encontradas operações de 'Entrada' (Compras/Fornecedores) nos arquivos selecionados.\n\n"
                    "O dashboard será exibido vazio.")
                self.df_suppliers_abc = pd.DataFrame()
            else:
                self.df_suppliers_abc = logic.calculate_supplier_abc(self.df_base)

            self.reset_all_filters()
        except Exception as e:
            QMessageBox.critical(self, "Erro ao Popular Dashboard", f"Ocorreu um erro inesperado ao processar os dados.\n\nDetalhes: {e}")
            traceback.print_exc()
            self.df_base = pd.DataFrame()
            self.df_suppliers_abc = pd.DataFrame()
            self.reset_all_filters()
        finally:
            self.is_populating = False

    def populate_supplier_table(self, suppliers_df):
        self.tree_suppliers.clear()
        if suppliers_df is None: return

        for _, row in suppliers_df.iterrows():
            valor_formatado = f"R$ {row['Valor_Total']:,.2f}"
            item = NumericTreeWidgetItem([
                "",
                row['Curva_ABC'],
                row['CNPJ/CPF Participante'],
                row['Nome Participante'],
                valor_formatado
            ])
            item.setData(2, Qt.ItemDataRole.UserRole, row['CNPJ/CPF Participante'])
            item.setData(4, Qt.ItemDataRole.UserRole, row['Valor_Total'])
            item.setTextAlignment(0, Qt.AlignmentFlag.AlignCenter)
            item.setTextAlignment(1, Qt.AlignmentFlag.AlignCenter)
            item.setTextAlignment(4, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.tree_suppliers.addTopLevelItem(item)
        self.tree_suppliers.sortByColumn(4, Qt.SortOrder.DescendingOrder)

    def update_ncm_donut_chart(self, data_df):
        """Atualiza apenas o gráfico NCM."""
        if self.ncm_donut_cid:
            try:
                self.fig_ncm_donut.canvas.mpl_disconnect(self.ncm_donut_cid)
            except Exception as e:
                print(f"Aviso: Falha ao desconectar sinal do gráfico NCM: {e}")
            self.ncm_donut_cid = None

        self.ax_ncm_donut.clear()
        if data_df.empty:
            self.ax_ncm_donut.text(0.5, 0.5, 'Sem dados', ha='center', va='center', color='grey')
            self.canvas_ncm_donut.draw()
            self.ncm_donut_cid = self.fig_ncm_donut.canvas.mpl_connect('pick_event', self.on_ncm_click)
            return

        ncm_summary = logic.calculate_ncm_summary(data_df)
        labels = ncm_summary['NCM']
        sizes = ncm_summary['Valor_Total']

        wedges, texts, autotexts = self.ax_ncm_donut.pie(
            sizes, labels=None, autopct='%1.1f%%', startangle=90,
            pctdistance=0.85, colors=STYLE_COLORS['charts'],
            wedgeprops=dict(width=0.4, edgecolor='w')
        )

        plt.setp(autotexts, size=8, weight="bold", color="white")

        for i, p in enumerate(wedges):
            ang = (p.theta2 - p.theta1)/2. + p.theta1
            y = np.sin(np.deg2rad(ang))
            x = np.cos(np.deg2rad(ang))
            horizontalalignment = {-1: "right", 1: "left"}[int(np.sign(x))]
            x_text = 1.2 * np.sign(x)
            y_text = 1.2 * y
            label_text = labels.iloc[i]

            p.set_label(label_text)
            p.set_picker(True)

            self.ax_ncm_donut.annotate(label_text, xy=(x, y), xytext=(x_text, y_text),
                                  horizontalalignment=horizontalalignment,
                                  arrowprops=dict(arrowstyle="-", color='gray', connectionstyle=f"angle,angleA=0,angleB={ang}"),
                                  bbox=dict(boxstyle="square,pad=0.3", fc="w", ec="k", lw=0.72))

        self.ax_ncm_donut.axis('equal')
        self.canvas_ncm_donut.draw()
        self.ncm_donut_cid = self.fig_ncm_donut.canvas.mpl_connect('pick_event', self.on_ncm_click)

    def update_cfop_donut_chart(self, data_df):
        """Atualiza o gráfico de rosca de CFOP."""
        if self.cfop_donut_cid:
            try:
                self.fig_cfop_donut.canvas.mpl_disconnect(self.cfop_donut_cid)
            except Exception as e:
                print(f"Aviso: Falha ao desconectar sinal do gráfico CFOP: {e}")
            self.cfop_donut_cid = None

        self.ax_cfop_donut.clear()
        if data_df.empty:
            self.ax_cfop_donut.text(0.5, 0.5, 'Sem dados', ha='center', va='center', color='grey')
            self.canvas_cfop_donut.draw()
            self.cfop_donut_cid = self.fig_cfop_donut.canvas.mpl_connect('pick_event', self.on_cfop_click)
            return

        cfop_summary = logic.calculate_cfop_summary(data_df)
        labels = cfop_summary['CFOP']
        sizes = cfop_summary['Valor_Total']

        wedges, texts, autotexts = self.ax_cfop_donut.pie(
            sizes, labels=None, autopct='%1.1f%%', startangle=90,
            pctdistance=0.85, colors=STYLE_COLORS['charts'],
            wedgeprops=dict(width=0.4, edgecolor='w')
        )

        plt.setp(autotexts, size=8, weight="bold", color="white")

        for i, p in enumerate(wedges):
            ang = (p.theta2 - p.theta1)/2. + p.theta1
            y = np.sin(np.deg2rad(ang))
            x = np.cos(np.deg2rad(ang))
            horizontalalignment = {-1: "right", 1: "left"}[int(np.sign(x))]
            x_text = 1.2 * np.sign(x)
            y_text = 1.2 * y
            label_text = labels.iloc[i]

            p.set_label(label_text)
            p.set_picker(True)

            self.ax_cfop_donut.annotate(label_text, xy=(x, y), xytext=(x_text, y_text),
                                  horizontalalignment=horizontalalignment,
                                  arrowprops=dict(arrowstyle="-", color='gray', connectionstyle=f"angle,angleA=0,angleB={ang}"),
                                  bbox=dict(boxstyle="square,pad=0.3", fc="w", ec="k", lw=0.72))

        self.ax_cfop_donut.axis('equal')
        self.canvas_cfop_donut.draw()
        self.cfop_donut_cid = self.fig_cfop_donut.canvas.mpl_connect('pick_event', self.on_cfop_click)

    def update_line_chart(self, data_df):
        self.ax_line.clear()
        if data_df.empty:
            self.ax_line.text(0.5, 0.5, 'Sem dados', ha='center', va='center', color='grey')
            self.canvas_line.draw()
            return

        monthly_summary = logic.calculate_monthly_summary(data_df)
        if monthly_summary.empty:
            self.ax_line.text(0.5, 0.5, 'Sem dados', ha='center', va='center', color='grey')
            self.canvas_line.draw()
            return

        self.ax_line.plot(monthly_summary['MesAno_Str'], monthly_summary['Valor_Total'],
                            marker='o', linestyle='-', color=STYLE_COLORS['primary_accent'])
        self.ax_line.grid(axis='y', color=STYLE_COLORS['grid'], linestyle='--')
        self.ax_line.spines['top'].set_visible(False)
        self.ax_line.spines['right'].set_visible(False)
        self.ax_line.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f"R$ {x:,.0f}"))
        plt.setp(self.ax_line.get_xticklabels(), rotation=30, horizontalalignment='right')
        self.fig_line.tight_layout()
        self.canvas_line.draw()

    # --- Funções de Filtro (Cross-filtering) ---

    def on_supplier_select(self, item, column):
        """Filtra os gráficos pelo Fornecedor selecionado."""
        try:
            if not isinstance(item, NumericTreeWidgetItem):
                 return
            supplier_id = item.data(2, Qt.ItemDataRole.UserRole)
            if not supplier_id: return

            if supplier_id == self.selected_supplier_cnpj:
                self.reset_all_filters()
                return

            self.selected_supplier_cnpj = supplier_id
            self.selected_ncm = None
            self.selected_cfop = None

            df_filtered = self.df_base[self.df_base['CNPJ/CPF Participante'] == self.selected_supplier_cnpj]

            self.update_ncm_donut_chart(df_filtered)
            self.update_cfop_donut_chart(df_filtered)
            self.update_line_chart(df_filtered)

        except Exception as e:
            QMessageBox.warning(self, "Erro de Seleção", f"Não foi possível aplicar o filtro: {e}")

    def on_ncm_click(self, event):
        """Filtra os gráficos e a tabela pelo NCM selecionado."""
        if self.is_populating:
            return
        try:
            if event is None or event.artist is None: return
            ncm_clicked = event.artist.get_label()
            if ncm_clicked is None: return

            if ncm_clicked == self.selected_ncm:
                self.reset_all_filters()
                return

            self.selected_ncm = ncm_clicked
            self.selected_supplier_cnpj = None
            self.selected_cfop = None

            self.tree_suppliers.clearSelection()

            df_filtered = self.df_base[self.df_base['NCM'] == self.selected_ncm]

            self.update_line_chart(df_filtered)
            self.update_cfop_donut_chart(df_filtered)

            df_suppliers_ncm = logic.calculate_supplier_abc(df_filtered)
            self.populate_supplier_table(df_suppliers_ncm)

        except Exception as e:
            print(f"--- ERRO FATAL em on_ncm_click ---")
            traceback.print_exc()
            print(f"----------------------------------")

    def on_cfop_click(self, event):
        """Filtra os gráficos e a tabela pelo CFOP selecionado."""
        if self.is_populating:
            return
        try:
            if event is None or event.artist is None: return
            cfop_clicked = event.artist.get_label()
            if cfop_clicked is None: return

            if cfop_clicked == self.selected_cfop:
                self.reset_all_filters()
                return

            self.selected_cfop = cfop_clicked
            self.selected_supplier_cnpj = None
            self.selected_ncm = None

            self.tree_suppliers.clearSelection()

            df_filtered = self.df_base[self.df_base['CFOP'] == self.selected_cfop]

            self.update_line_chart(df_filtered)
            self.update_ncm_donut_chart(df_filtered)

            df_suppliers_cfop = logic.calculate_supplier_abc(df_filtered)
            self.populate_supplier_table(df_suppliers_cfop)

        except Exception as e:
            print(f"--- ERRO FATAL em on_cfop_click ---")
            traceback.print_exc()
            print(f"-----------------------------------")

    def reset_all_filters(self):
        """Reseta todos os filtros e atualiza todos os widgets."""
        self.selected_supplier_cnpj = None
        self.selected_ncm = None
        self.selected_cfop = None

        self.tree_suppliers.clearSelection()

        # Verifica se df_suppliers_abc existe e não está vazio antes de popular
        if hasattr(self, 'df_suppliers_abc') and not self.df_suppliers_abc.empty:
            self.populate_supplier_table(self.df_suppliers_abc)
        else:
             self.tree_suppliers.clear() # Limpa se não houver dados

        # Verifica se df_base existe antes de atualizar gráficos
        if hasattr(self, 'df_base'):
            self.update_ncm_donut_chart(self.df_base)
            self.update_cfop_donut_chart(self.df_base)
            self.update_line_chart(self.df_base)
        else:
            # Limpa os gráficos se não houver df_base
            self.update_ncm_donut_chart(pd.DataFrame())
            self.update_cfop_donut_chart(pd.DataFrame())
            self.update_line_chart(pd.DataFrame())


# ==============================================================================
# CLASSE PRINCIPAL DA APLICAÇÃO (Gerenciador de Páginas)
# ==============================================================================

class AnaliseFiscalApp(QWidget):
    """
    Esta é agora a janela principal, que gerencia o QStackedWidget
    para alternar entre o Menu e o Dashboard.
    """
    window_closed = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Análise EFD Contribuições - Bloco C")
        self.setStyleSheet(STYLESHEET)

        self.stack = QStackedWidget(self)

        self.menu_widget = MenuWidget()
        self.menu_widget.data_processed.connect(self.show_dashboard)

        self.dashboard_widget = DashboardWidget()
        self.dashboard_widget.back_to_menu_requested.connect(self.show_menu)

        self.stack.addWidget(self.menu_widget)
        self.stack.addWidget(self.dashboard_widget)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.addWidget(self.stack)

        self.stack.setCurrentWidget(self.menu_widget)
        self.setGeometry(200, 200, 750, 550)

    def show_dashboard(self, df_raw):
        """
        Slot que é ativado quando o MenuWidget emite o sinal 'data_processed'.
        """
        try:
            self.stack.setCurrentWidget(self.dashboard_widget)
            self.resize(1500, 900)
            self.showMaximized()
            QApplication.processEvents()
            self.dashboard_widget.populate_dashboard(df_raw)

        except Exception as e:
            print(f"\n--- ERRO FATAL em show_dashboard ---")
            traceback.print_exc()
            print(f"----------------------------------\n")
            QMessageBox.critical(self, "Erro ao Exibir Dashboard",
                f"Ocorreu um erro ao tentar exibir os gráficos:\n\n{e}")
            self.stack.setCurrentWidget(self.menu_widget)


    def show_menu(self):
        """Muda o QStackedWidget de volta para a página do menu."""
        self.showNormal()
        self.resize(750, 550)
        self.stack.setCurrentWidget(self.menu_widget)

    def closeEvent(self, event):
        """Emite um sinal quando a janela é fechada."""
        self.window_closed.emit()
        super().closeEvent(event)

# --- Ponto de Entrada (main) ---
def main():
    """Função principal para ser chamada pelo Hub."""
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)

    window = AnaliseFiscalApp()
    window.show()
    return app, window

# Bloco para testar o aplicativo isoladamente
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = AnaliseFiscalApp()
    main_window.show()
    sys.exit(app.exec())