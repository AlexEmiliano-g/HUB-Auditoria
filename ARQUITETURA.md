# Documentação de Arquitetura - HUB Auditoria

## 📋 Visão Geral

O **HUB Auditoria** é uma aplicação desktop modular desenvolvida em Python que centraliza ferramentas de automação para as áreas de **Auditoria e Contabilidade**. A aplicação segue um padrão de arquitetura baseado em um **hub central launcher** que gerencia múltiplos módulos temáticos organizados por setor.

### Informações Técnicas
- **Linguagem**: Python 3.x
- **Framework GUI**: PyQt6
- **Propósito**: Automação de processos contábeis e fiscais
- **Tipo**: Aplicação Desktop (Standalone)
- **Sistema Operacional**: Windows

---

## 🏗️ Arquitetura Geral

```
┌─────────────────────────────────────────────────────────────┐
│                      HUB Central (hub.py)                   │
│  - Splash Screen (Loading)                                  │
│  - Menu Lateral com Setores                                 │
│  - Lista de Programas por Setor                             │
│  - Log de Execução                                          │
└────────────────────┬────────────────────────────────────────┘
                     │
         ┌───────────┼───────────┬──────────────┐
         │           │           │              │
    ┌────▼────┐  ┌───▼───┐  ┌───▼────┐  ┌─────▼─────┐
    │Contabil │  │Fiscal │  │Análise │  │ IA (Beta) │
    │   idade │  │       │  │        │  │           │
    └────┬────┘  └───┬───┘  └───┬────┘  └─────┬─────┘
         │           │          │              │
    ┌────▼────┐  ┌───▼───────────┴──────────────────────┐
    │Tabulador│  │  Relatórios Fiscais (8 Tipos)       │
    │         │  │  - C100 (EFD-C, SPED-F)             │
    │         │  │  - C500 (EFD-C, SPED-F)             │
    │         │  │  - CIAP (SPED-F)                    │
    │         │  │  - CT (Controladoria)               │
    │         │  │  - D100 (SPED-F)                    │
    │         │  │  - F100 (EFD-C)                     │
    │         │  │  - XML NF-e                         │
    │         │  └────┬────────────────────────────────┘
    └──────────┘      │
                  ┌───▼───────┐
                  │ Análise    │
                  │ Fiscal     │
                  └────────────┘
```

---

## 📁 Estrutura de Diretórios

```
HUB-Auditoria/
├── hub.py                          # Aplicação Principal (Hub Launcher)
├── requirements.txt                # Dependências do Projeto
├── README.md                        # Documentação Inicial
├── ARQUITETURA.md                  # Este arquivo
├── DEPENDENCIAS.md                 # Documento de Dependências (gerado)
├── icone.ico                        # Ícone da Aplicação
├── logo.png                         # Logo Principal do HUB
├── logo_aber.png                    # Logo de Abertura (Splash)
│
└── scripts/                         # Diretório Principal de Módulos
    ├── Contabilidade/
    │   ├── tabulador_app.py         # Interface GUI do Tabulador
    │   ├── tabulador_logic.py       # Lógica de Processamento
    │   └── __pycache__/
    │
    ├── Fiscal/
    │   ├── relatorio_C100_EFD_C_app.py
    │   ├── relatorio_C100_SPED_F_app.py
    │   ├── relatorio_C500_EFD_C_app.py
    │   ├── relatorio_C500_SPED_F_app.py
    │   ├── relatorio_CIAP_SPED_F_app.py
    │   ├── relatorio_CT_app.py
    │   ├── relatorio_D100_SPED_F_app.py
    │   ├── relatorio_F100_EFD_C_app.py
    │   ├── relatorio_XML_app.py
    │   └── __pycache__/
    │
    ├── Analise/
    │   ├── analise_fiscal_app.py    # Interface GUI de Análise
    │   ├── analise_fiscal_logic.py  # Lógica de Análise
    │   └── __pycache__/
    │
    └── IA/
        ├── Texto_anonimo_app.py     # Módulo de Anonimização com IA (Beta)
        └── __pycache__/
```

---

## 🔄 Fluxo de Execução

### 1. Inicialização da Aplicação

```
┌──────────────────────────────────┐
│ Execução: python hub.py          │
└────────────┬─────────────────────┘
             │
             ▼
┌──────────────────────────────────┐
│ Splash Screen (Fade In)          │
│ Status: "Inicializando..."       │
└────────────┬─────────────────────┘
             │
             ▼
┌──────────────────────────────────┐
│ start_application(app, splash)   │
│ - Importa pandas                 │
│ - Importa openpyxl               │
│ - Inicializa Hub()               │
└────────────┬─────────────────────┘
             │
             ▼
┌──────────────────────────────────┐
│ Hub Window (Fade Out Splash)     │
│ - Menu Lateral (Setores)         │
│ - Área de Programas              │
│ - Log de Execução                │
└──────────────────────────────────┘
```

### 2. Seleção e Execução de Módulo

```
Usuário Seleciona Setor
    │
    ├─▶ Hub.load_scripts(setor)
    │   └─▶ Escaneia pasta: scripts/{setor}/
    │       └─▶ Identifica arquivos *_app.py
    │           └─▶ Popula lista de programas
    │
Usuário Seleciona Programa
    │
    └─▶ Hub.run_script()
        ├─▶ Importa módulo dinamicamente
        │   module_path = f"scripts.{setor}.{module_name}"
        │
        ├─▶ Executa function main()
        │   └─▶ Retorna (app, window)
        │
        └─▶ Exibe janela do programa
            └─▶ Log de Status
```

---

## 🧩 Componentes Principais

### 1. **Hub Central** (`hub.py`)

**Responsabilidades:**
- Gerenciamento da interface principal
- Carregamento dinâmico de módulos
- Controle do ciclo de vida da aplicação
- Renderização de Splash Screen

**Classes:**
- `Splash`: Janela de boas-vindas com animação de fade
- `Hub`: Janela principal com menu de setores e programas

**Métodos Principais:**
```python
Hub.load_scripts(item)      # Carrega scripts de um setor
Hub.run_script()            # Executa script selecionado
start_application(app, splash)  # Inicializa aplicação com loading
```

**Padrão Utilizado:** Lazy Loading + Dynamic Module Import

---

### 2. **Setor: Contabilidade**

#### Tabulador de Balancetes (`tabulador_app.py` + `tabulador_logic.py`)

**Objetivo:** Processar e tabular múltiplos arquivos de balancete de diferentes sistemas.

**Funcionalidades:**
- Seleção de múltiplos arquivos Excel
- Suporte a diferentes sistemas: Fuga, Uniair
- Processamento em batch
- Geração de relatórios consolidados

**Tecnologias:**
- pandas: Processamento de dados
- openpyxl: Manipulação de Excel
- PyQt6: Interface gráfica

**Fluxo:**
```
Selecionar Arquivos (CSV/XLSX)
    │
    ▼
Escolher Sistema (Fuga/Uniair)
    │
    ▼
processar_arquivos_selecionados()
    ├─ Lê cada arquivo
    ├─ Padroniza dados
    ├─ Consolida tabelas
    │
    ▼
Exportar Resultado (XLSX)
```

---

### 3. **Setor: Fiscal**

#### Módulos de Relatórios Fiscais (8 Tipos)

**Objetivo:** Gerar e analisar relatórios fiscais de diferentes formatos EFD-C e SPED-F.

| Módulo | Tipo | Formato | Descrição |
|--------|------|---------|-----------|  
| `relatorio_C100_EFD_C_app.py` | C100 | EFD-C | Registro de Operações com ICMS |
| `relatorio_C100_SPED_F_app.py` | C100 | SPED-F | Registro de Operações com ICMS (Alimentar) |
| `relatorio_C500_EFD_C_app.py` | C500 | EFD-C | Registro de Consolidação |
| `relatorio_C500_SPED_F_app.py` | C500 | SPED-F | Registro de Consolidação (Alimentar) |
| `relatorio_CIAP_SPED_F_app.py` | CIAP | SPED-F | Controle de ICMS a Apropriação |
| `relatorio_CT_app.py` | CT | Controladoria | Relatórios de Controladoria |
| `relatorio_D100_SPED_F_app.py` | D100 | SPED-F | Documentos Fiscais (NFe, CTE) |
| `relatorio_F100_EFD_C_app.py` | F100 | EFD-C | Operações Internas |
| `relatorio_XML_app.py` | XML | NF-e | Análise de Nota Fiscal Eletrônica |

---

### 4. **Setor: Análise**

#### Análise Fiscal (`analise_fiscal_app.py` + `analise_fiscal_logic.py`)

**Objetivo:** Análise comparativa e visualização de dados fiscais com gráficos interativos.

**Funcionalidades:**
- Carregamento de múltiplos arquivos de dados
- Análise comparativa entre períodos/entidades
- Visualização em gráficos (Matplotlib + PyQt6)
- Exportação de análises em PDF
- Interface com árvore de dados (TreeWidget)

---

### 5. **Setor: IA** (Beta)

#### Anonimização de Texto (`Texto_anonimo_app.py`)

**Objetivo:** Processamento de textos com capacidades de anonimização baseadas em IA.

**Status:** Beta (Desenvolvimento)

---

## 🔌 Dependências do Projeto

### Dependências Python

| Biblioteca | Versão | Propósito | Setor |
|-----------|--------|----------|-------|
| `pandas` | Latest | Processamento de dados | Contabilidade, Fiscal, Análise |
| `PyQt6` | Latest | Framework GUI | Todos |
| `openpyxl` | Latest | Leitura/escrita Excel | Contabilidade, Fiscal, Análise |
| `numpy` | Latest | Operações numéricas | Análise |
| `matplotlib` | Latest | Geração de gráficos | Análise |

---

## 🎯 Padrões de Design Implementados

### 1. **Lazy Loading**
- Módulos são importados apenas quando necessário
- Reduz tempo de inicialização

### 2. **Factory Pattern**
- Cada módulo expõe uma função `main()` que retorna a aplicação e janela

### 3. **Separação de Responsabilidades**
- Padrão MVC Light:
  - `*_app.py`: View (Interface GUI com PyQt6)
  - `*_logic.py`: Model/Logic (Processamento de dados)

---

## 📈 Performance e Otimizações

### Implementadas

1. **Lazy Loading de Módulos**
   - Pandas e openpyxl importados apenas no `start_application()`
   - Reduz tempo de inicialização do Hub

2. **Cache de Recursos**
   - Logos e ícones carregados uma única vez
   - Armazenados em memória durante execução

3. **Processamento em Batch**
   - Tabulador processa múltiplos arquivos sequencialmente
   - Consolidação final em um arquivo

---

**Documento Gerado:** 19/05/2026  
**Status:** Documentação Completa - Versão 1.0  
**Maintainer:** Tim do HUB Auditoria