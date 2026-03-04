<div align="center">

<img src="icone.png" width="120" alt="Sorteador Logo"/>

# 🎲 Sorteador de Base de Dados

**Sorteio auditável, rastreável e reprodutível a partir de planilhas Excel**

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![pandas](https://img.shields.io/badge/pandas-2.x-150458?style=for-the-badge&logo=pandas&logoColor=white)](https://pandas.pydata.org)
[![openpyxl](https://img.shields.io/badge/openpyxl-3.x-217346?style=for-the-badge)](https://openpyxl.readthedocs.io)
[![License](https://img.shields.io/badge/Licença-MIT-green?style=for-the-badge)](LICENSE)

[Funcionalidades](#-funcionalidades) • [Instalação](#-instalação) • [Como Usar](#-como-usar) • [Rastreabilidade](#-rastreabilidade) • [Arquivos Gerados](#-arquivos-gerados)

</div>

---

## 📋 Sobre o Projeto

O **Sorteador de Base de Dados** é uma aplicação desktop desenvolvida em Python que permite realizar sorteios a partir de planilhas Excel de forma **justa, auditável e reprodutível**.

Ideal para sorteios corporativos, fiscalizações, auditorias e qualquer processo que exija conformidade e rastreabilidade completa — cada sorteio gera evidências que permitem verificação e reprodução posterior.

---

## ✨ Funcionalidades

- 📂 **Importação direta** de planilhas `.xlsx` e `.xls`
- 🔍 **Filtro por coluna denominadora** — sorteia apenas dentro de um segmento específico
- 🎯 **Sorteio sem reposição** — cada registro só pode ser selecionado uma vez
- 🔒 **Hash SHA-256** da base importada — prova de integridade dos dados
- 🔑 **Semente de sorteio** — permite reproduzir qualquer sorteio identicamente
- 📄 **Relatório `.txt`** gerado automaticamente após cada sorteio
- 📊 **Dois arquivos Excel** de resultado: base completa e apenas sorteados
- 🗂️ **Log JSON** com histórico completo de todos os sorteios realizados
- 🖥️ **Interface gráfica** intuitiva, sem necessidade de conhecimento técnico

---

## 🚀 Instalação

### Pré-requisitos

- [Python 3.10+](https://python.org/downloads) — marque **"Add Python to PATH"** durante a instalação

### Opção 1 — Gerar o `.exe` (recomendado para uso no dia a dia)

1. Clone ou baixe este repositório
2. Coloque todos os arquivos na mesma pasta:
   ```
   📁 pasta/
   ├── sorteador.py
   ├── build.bat
   ├── icone.ico
   └── icone.png
   ```
3. Clique duas vezes em **`build.bat`**
4. Aguarde 2–3 minutos — o executável será gerado em `dist/Sorteador_BaseDados.exe`

> O `.exe` gerado funciona em qualquer máquina Windows, sem precisar do Python instalado.

### Opção 2 — Rodar direto pelo Python

```bash
# Instale as dependências
pip install pandas openpyxl xlrd pillow

# Execute
python sorteador.py
```

---

## 📖 Como Usar

### 1️⃣ Importar Planilha

Clique em **"Selecionar arquivo Excel"** e escolha seu `.xlsx` ou `.xls`.

O programa exibirá o hash SHA-256 da base — este código é a impressão digital da sua planilha e comprova que ela não foi alterada.

### 2️⃣ Configurar o Sorteio

| Campo | Descrição |
|-------|-----------|
| **Coluna denominadora** | Coluna usada para segmentar o sorteio (ex: `Região`, `Tipo`, `Status`) |
| **Filtrar por valor** | Valor específico dessa coluna (ex: `Norte`, `Ativo`, `Categoria A`) |
| **Qtd. a sortear** | Quantidade de pontos a selecionar |
| **Semente** | Deixe em branco para sorteio automático — use um número fixo para reproduzir um sorteio anterior |

### 3️⃣ Sortear

Clique em **"▶ SORTEAR AGORA"**, escolha a pasta de destino e pronto.

---

## 📁 Arquivos Gerados

A cada sorteio, três arquivos são criados automaticamente:

```
📁 pasta_destino/
├── sorteio_completo_20260304_143522.xlsx       # Base inteira + coluna SORTEADO marcada
├── sorteio_selecionados_20260304_143522.xlsx   # Apenas os registros sorteados
└── relatorio_sorteio_20260304_143522.txt       # Relatório completo do sorteio
```

### Exemplo do relatório `.txt`

```
============================================================
  RELATÓRIO DE SORTEIO
============================================================

  Data e hora        : 04/03/2026 14:35:22

------------------------------------------------------------
  BASE DE DADOS
------------------------------------------------------------
  Arquivo            : C:\planilhas\base_2026.xlsx
  SHA-256 da base    : a3f9c2d1e8b74f6c...

------------------------------------------------------------
  CONFIGURAÇÃO DO SORTEIO
------------------------------------------------------------
  Coluna filtro      : Região
  Valor filtro       : Norte
  Total disponível   : 850 registros
  Quantidade sorteada: 20 pontos

------------------------------------------------------------
  IDENTIFICAÇÃO DO SORTEIO
------------------------------------------------------------
  Semente            : 2847361092

------------------------------------------------------------
  LOG DE AUDITORIA
------------------------------------------------------------
  C:\Users\usuario\Sorteador_Logs\historico_sorteios.json

============================================================
```

---

## 🔒 Rastreabilidade

Cada sorteio é registrado automaticamente em:

```
C:\Users\<usuario>\Sorteador_Logs\historico_sorteios.json
```

O log contém:

```json
{
  "data_hora": "2026-03-04T14:35:22",
  "arquivo_origem": "C:\\planilhas\\base_2026.xlsx",
  "hash_base_sha256": "a3f9c2d1e8b74f6c...",
  "coluna_filtro": "Região",
  "valor_filtro": "Norte",
  "total_disponivel": 850,
  "qtd_sorteada": 20,
  "semente": 2847361092,
  "indices_sorteados": [12, 47, 93, 104, ...],
  "arquivo_completo": "C:\\resultados\\sorteio_completo_20260304_143522.xlsx",
  "arquivo_sorteados": "C:\\resultados\\sorteio_selecionados_20260304_143522.xlsx"
}
```

### Como reproduzir um sorteio anterior

1. Abra o `historico_sorteios.json` e localize a semente do sorteio desejado
2. Importe a **mesma planilha original** (o hash SHA-256 confirma que é a mesma base)
3. Aplique o **mesmo filtro** (coluna + valor)
4. Informe a **semente** no campo correspondente
5. Clique em Sortear — o resultado será **100% idêntico**

> A combinação de **hash SHA-256 + semente** garante que qualquer auditor possa verificar e reproduzir qualquer sorteio realizado.

---

## 🛠️ Tecnologias

| Biblioteca | Uso |
|------------|-----|
| `tkinter` | Interface gráfica |
| `pandas` | Leitura e manipulação da planilha |
| `openpyxl` | Engine de leitura/escrita `.xlsx` |
| `xlrd` | Engine de leitura `.xls` (legado) |
| `Pillow` | Ícone com transparência |
| `hashlib` | Hash SHA-256 da base |
| `random.Random` | Sorteio com semente reprodutível |
| `PyInstaller` | Compilação para `.exe` |

---

## 🗂️ Estrutura do Repositório

```
📁 sorteador-base-dados/
├── sorteador.py          # Código-fonte principal
├── build.bat             # Script de compilação para .exe
├── icone.png             # Ícone do programa (PNG com transparência)
├── icone.ico             # Ícone para barra de tarefas Windows
├── README.md             # Este arquivo
└── README_Sorteador.docx # Manual de instruções completo
```

---

## ❓ Solução de Problemas

| Erro | Solução |
|------|---------|
| `Import openpyxl failed` | Execute o `build.bat` novamente — ele instala o `openpyxl` automaticamente |
| `Python não encontrado` | Reinstale o Python marcando **"Add Python to PATH"** |
| Planilha vazia | Verifique se a planilha tem dados na aba principal |
| Qtd. maior que disponível | Reduza a quantidade ou revise o filtro selecionado |
| Ícone com fundo preto | Use os arquivos `icone.ico` e `icone.png` do repositório |

---

## 👤 Autor

<div align="center">

**Pablo Bernar**

[![LinkedIn](https://img.shields.io/badge/LinkedIn-pablo--bernar-0A66C2?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/pablo-bernar/)

</div>
