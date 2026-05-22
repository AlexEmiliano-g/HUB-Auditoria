# Guia de Referência Rápida - HUB Auditoria

## 🎯 Início Rápido

### Setup Inicial

```bash
# 1. Clonar repositório
git clone https://github.com/AlexEmiliano-g/HUB-Auditoria.git
cd HUB-Auditoria

# 2. Criar ambiente virtual
python -m venv venv
venv\Scripts\activate  # Windows

# 3. Instalar dependências
pip install -r requirements.txt

# 4. Executar aplicação
python hub.py
```

---

## 📊 Arquitetura em Níveis

### Nível 1: Camada de Aplicação
```
┌──────────────────────────────┐
│ hub.py - Hub Central         │
│ (Launcher + Router)          │
└──────────────────────────────┘
```

### Nível 2: Camada de Setores
```
┌─────────────────────────────────────────────────────────────┐
│  Contabilidade  │  Fiscal  │  Análise  │  IA (Beta)       │
└─────────────────────────────────────────────────────────────┘
```

### Nível 3: Camada de Módulos
```
Contabilidade:
  ├─ Tabulador
  └─ [Interface + Lógica]

Fiscal:
  ├─ 8 Relatórios Específicos
  ├─ XML Handler
  └─ [Processamento + UI]

Análise:
  ├─ Análise Fiscal
  ├─ Gráficos Interativos
  └─ [Visualização + Cálculos]

IA:
  └─ Anonimização de Texto
```

### Nível 4: Camada de Dependências
```
PyQt6 (GUI)
  ├─ Widgets
  ├─ Signals/Slots
  └─ Stylesheets

pandas (Data Processing)
  ├─ DataFrame
  ├─ I/O (Excel, CSV)
  └─ Transformations

openpyxl (Excel)
  ├─ Read/Write
  └─ Formatting

matplotlib (Visualization)
  ├─ Charts
  └─ PDF Export
```

---

## 📂 Estrutura de Diretórios - Árvore Completa

```
HUB-Auditoria/
│
├── 📄 hub.py                        # ⭐ Aplicação Principal
├── 📄 requirements.txt              # Dependências Python
├── 📄 README.md                     # Documentação Inicial
├── 📄 ARQUITETURA.md               # 📚 Este projeto (detalhado)
├── 📄 DEPENDENCIAS.md              # 📦 Gerenciamento de deps
├── 📄 REFERENCIA_RAPIDA.md         # 🚀 Este arquivo
│
├── 🖼️  icone.ico                    # Ícone da Aplicação
├── 🖼️  logo.png                     # Logo Principal
├── 🖼️  logo_aber.png                # Logo de Splash
│
├── 📁 scripts/                      # Centro de Módulos
│   │
│   ├── 📁 Contabilidade/
│   │   ├── tabulador_app.py         # [UI] Tabulador
│   │   ├── tabulador_logic.py       # [LOGIC] Processamento
│   │   └── __pycache__/
│   │
│   ├── 📁 Fiscal/
│   │   ├── relatorio_C100_EFD_C_app.py
│   │   ├── relatorio_C100_SPED_F_app.py
│   │   ├── relatorio_C500_EFD_C_app.py
│   │   ├── relatorio_C500_SPED_F_app.py
│   │   ├── relatorio_CIAP_SPED_F_app.py
│   │   ├── relatorio_CT_app.py
│   │   ├── relatorio_D100_SPED_F_app.py
│   │   ├── relatorio_F100_EFD_C_app.py
│   │   ├── relatorio_XML_app.py     # Handler XML NF-e
│   │   └── __pycache__/
│   │
│   ├── 📁 Analise/
│   │   ├── analise_fiscal_app.py    # [UI] Com gráficos
│   │   ├── analise_fiscal_logic.py  # [LOGIC] Processamento
│   │   └── __pycache__/
│   │
│   └── 📁 IA/
│       ├── Texto_anonimo_app.py     # [BETA] Anonimização
│       └── __pycache__/
│
└── 📁 .git/                         # Controle de versão
```

---

## 🔧 Fluxo de Execução Simplificado

### Inicialização

```
┌─────────────────────┐
│ python hub.py       │ ← Inicia processo
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│ QApplication(argv)  │ ← Cria app Qt
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│ Splash Screen       │ ← Loading visual
│ (2 segundos)        │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│ start_application() │ ← Carrega libs pesadas
│ - pandas            │
│ - openpyxl          │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│ Hub Window (Main)   │ ← Interface principal
└─────────────────────┘
```

### Execução de Módulo

```
Usuário Clica Setor
    │
    ▼
load_scripts(setor)
    │
    ├─ Verifica: scripts/{setor}/
    │
    ▼
Popula QListWidget com *_app.py
    │
    ▼
Usuário Seleciona Programa
    │
    ▼
run_script()
    │
    ├─ module_path = f"scripts.{setor}.{modulo}"
    │
    ├─ importlib.import_module(module_path)
    │
    ├─ modulo.main() ← Retorna (app, window)
    │
    ▼
Exibe Janela do Programa
    │
    └─ Log de Status atualizado
```

---

## 🎨 Paleta de Cores Padrão

| Uso | Cor | Hex | RGB |
|-----|-----|-----|-----|
| Primária | Azul Google | #4285F4 | 66, 133, 244 |
| Secundária | Azul Escuro | #1967D2 | 25, 103, 210 |
| Hover | Azul Mais Escuro | #1A73E8 | 26, 115, 232 |
| Pressionado | Azul Muito Escuro | #1765CC | 23, 101, 204 |
| Fundo | Cinza Claro | #F8F9FA | 248, 249, 250 |
| Borda | Cinza Linha | #E0E0E0 | 224, 224, 224 |
| Seleção | Azul Fundo | #E8F0FE | 232, 240, 254 |
| Texto | Cinza Escuro | #3C4043 | 60, 64, 67 |
| Texto Fundo | Branco | #FFFFFF | 255, 255, 255 |

---

## 📋 Padrões de Implementação

### Padrão 1: Novo Módulo de Aplicação

```python
# novo_modulo_app.py
import sys
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel
from PyQt6.QtGui import QFont

class NovoModuloApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
    
    def initUI(self):
        self.setWindowTitle("Novo Módulo")
        self.resize(600, 400)
        
        layout = QVBoxLayout(self)
        
        titulo = QLabel("Bem-vindo ao Novo Módulo")
        titulo.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        layout.addWidget(titulo)
        
        botao = QPushButton("Executar")
        botao.clicked.connect(self.executar)
        layout.addWidget(botao)
    
    def executar(self):
        print("Executando...")
    
def main():
    app = QApplication(sys.argv)
    window = NovoModuloApp()
    window.show()
    return (app, window)
```

### Padrão 2: Processamento com pandas

```python
# novo_modulo_logic.py
import pandas as pd

def processar_dados(arquivo_path):
    # Ler dados
    df = pd.read_excel(arquivo_path)
    
    # Transformar
    df['nova_coluna'] = df['coluna_existente'].apply(lambda x: x * 2)
    
    # Consolidar
    resultado = df.groupby('categoria').sum()
    
    return resultado

def exportar_excel(df, output_path):
    df.to_excel(output_path, index=True)
```

### Padrão 3: Tratamento de XML

```python
import xml.etree.ElementTree as ET

NAMESPACE = {'ns': 'http://exemplo.com'}

def safe_get(element, path, default=''):
    try:
        found = element.find(path, NAMESPACE)
        return found.text.strip() if found is not None else default
    except:
        return default

def processar_xml(arquivo_xml):
    tree = ET.parse(arquivo_xml)
    root = tree.getroot()
    
    for item in root.findall('ns:item', NAMESPACE):
        valor = safe_get(item, 'ns:valor')
        print(f"Valor: {valor}")
```

---

## 🚨 Troubleshooting Comum

### Problema 1: Módulo não carrega
```
ERRO de Importação: Não foi possível carregar o módulo 'scripts.Contabilidade.tabulador_app'
```

**Solução:**
- ✅ Verificar se arquivo termina em `_app.py`
- ✅ Verificar se pasta existe em `scripts/{SETOR}/`
- ✅ Verificar se função `main()` existe no módulo
- ✅ Testar import manualmente: `python -c "from scripts.Contabilidade import tabulador_app"`

### Problema 2: pandas não encontrado
```
ModuleNotFoundError: No module named 'pandas'
```

**Solução:**
```bash
pip install pandas
# ou
pip install -r requirements.txt
```

### Problema 3: Arquivo Excel corrompido
```
openpyxl.utils.exceptions.InvalidFileException: File is not a zip
```

**Solução:**
1. Abrir arquivo no Excel
2. Salvar como novo arquivo
3. Tentar novamente

### Problema 4: XML namespace não encontra elementos
```
safe_get() retorna sempre default
```

**Solução:**
- ✅ Verificar namespace correto
- ✅ Printar XML para debug: `print(ET.tostring(element))`
- ✅ Usar ferramenta como XMLSpy para validar

---

## 📊 Tabela de Módulos Disponíveis

### Contabilidade (1 módulo)
| Módulo | Status | Entrada | Saída |
|--------|--------|---------|-------|
| Tabulador | ✅ Ativo | XLSX/CSV (múltiplos) | XLSX Consolidado |

### Fiscal (9 módulos)
| Módulo | Status | Entrada | Saída |
|--------|--------|---------|-------|
| C100 EFD-C | ✅ Ativo | EFD-C | XLSX Relatório |
| C100 SPED-F | ✅ Ativo | SPED-F | XLSX Relatório |
| C500 EFD-C | ✅ Ativo | EFD-C | XLSX Relatório |
| C500 SPED-F | ✅ Ativo | SPED-F | XLSX Relatório |
| CIAP SPED-F | ✅ Ativo | SPED-F | XLSX Relatório |
| CT | ✅ Ativo | Dados CT | XLSX Relatório |
| D100 SPED-F | ✅ Ativo | SPED-F | XLSX Relatório |
| F100 EFD-C | ✅ Ativo | EFD-C | XLSX Relatório |
| XML NF-e | ✅ Ativo | XML | XLSX Relatório |

### Análise (1 módulo)
| Módulo | Status | Entrada | Saída |
|--------|--------|---------|-------|
| Análise Fiscal | ✅ Ativo | XLSX (múltiplos) | Gráficos + PDF |

### IA (1 módulo - Beta)
| Módulo | Status | Entrada | Saída |
|--------|--------|---------|-------|
| Anonimização | 🔄 Beta | Texto | Texto Anonimizado |

---

## 🔐 Segurança e Boas Práticas

### ✅ Implementado
- Validação de arquivo (try-except)
- Conversão segura de tipos
- Tratamento de namespaces XML
- Isolamento de módulos

### 📝 Recomendações
- [ ] Adicionar logging estruturado
- [ ] Validar entrada do usuário
- [ ] Implementar testes automatizados
- [ ] Adicionar auditoria de mudanças
- [ ] Criptografar dados sensíveis

---

## 📈 Performance

### Tempos Típicos

| Operação | Tempo |
|----------|-------|
| Inicialização (com Splash) | ~2-3s |
| Carregamento de lista de programas | ~100ms |
| Leitura de arquivo Excel (1MB) | ~500ms |
| Processamento de 10 arquivos | ~2-5s |
| Geração de gráficos | ~1-2s |
| Exportação para PDF | ~500ms-1s |

---

## 🔄 Workflow Típico do Usuário

### 1. Usuário Contábil

```
1. Executar hub.py
2. Selecionar "Contabilidade"
3. Selecionar "Tabulador"
4. Procurar arquivos de balancete
5. Escolher sistema (Fuga/Uniair)
6. Clicar "Iniciar Tabulação"
7. Salvar arquivo consolidado
```

### 2. Usuário Fiscal

```
1. Executar hub.py
2. Selecionar "Fiscal"
3. Selecionar relatório (ex: XML NF-e)
4. Selecionar arquivo fiscal
5. Clicar "Processar"
6. Revisar dados extraídos
7. Exportar para Excel
```

### 3. Usuário Analista

```
1. Executar hub.py
2. Selecionar "Análise"
3. Selecionar "Análise Fiscal"
4. Carregar arquivo de dados
5. Escolher período/comparação
6. Visualizar gráficos
7. Exportar análise em PDF
```

---

## 📞 Suporte Rápido

### Contatos por Área

| Setor | Responsável | Contato |
|-------|-------------|---------|
| Arquitetura | Equipe HUB | hub-team@empresa.com |
| Contabilidade | Especialista Cont. | contab@empresa.com |
| Fiscal | Especialista Fiscal | fiscal@empresa.com |
| Análise | Analista | analise@empresa.com |

### Recursos

- 📖 [ARQUITETURA.md](./ARQUITETURA.md) - Documentação completa
- 📦 [DEPENDENCIAS.md](./DEPENDENCIAS.md) - Gerenciamento de dependências
- 🐛 [GitHub Issues](https://github.com/AlexEmiliano-g/HUB-Auditoria/issues)
- 📧 Email técnico: tech@empresa.com

---

## 🎓 Guias de Aprendizado

### Iniciante
1. Ler [README.md](./README.md)
2. Executar `python hub.py`
3. Explorar módulos da interface
4. Revisar [ARQUITETURA.md](./ARQUITETURA.md) - Seção "Componentes Principais"

### Intermediário
1. Revisar código-fonte de um módulo simples (tabulador)
2. Estudar padrão MVC em `*_app.py` + `*_logic.py`
3. Implementar pequeno módulo de teste
4. Fazer commit e PR no repositório

### Avançado
1. Implementar novo tipo de relatório fiscal
2. Otimizar performance de processamento
3. Adicionar testes automatizados
4. Revisar segurança e boas práticas

---

## ✨ Quick Commands

```bash
# Setup
python -m venv venv && venv\Scripts\activate && pip install -r requirements.txt

# Executar
python hub.py

# Testar imports
python -c "import pandas; import PyQt6; import openpyxl; print('OK')"

# Verificar dependências
pip list

# Atualizar dependências
pip install --upgrade -r requirements.txt

# Gerar novo requirements.txt
pip freeze > requirements.txt

# Limpar cache
del /s __pycache__

# Git commands
git status
git add .
git commit -m "Mensagem"
git push origin main
```

---

## 📚 Recursos Externos

- Python: https://python.org
- PyQt6: https://www.riverbankcomputing.com/software/pyqt
- pandas: https://pandas.pydata.org
- openpyxl: https://openpyxl.readthedocs.io
- matplotlib: https://matplotlib.org

---

**Última Atualização:** 19/05/2026  
**Versão:** 1.0  
**Status:** Pronto para Uso
