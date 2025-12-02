# 📊 Excel Consolidation Script (Python/Pandas)

Este repositório contém um script em Python que automatiza a consolidação de dados de múltiplos arquivos Excel (`.xlsx`) localizados em uma única pasta, combinando-os em um único arquivo de saída.

O script utiliza a poderosa biblioteca **Pandas** para uma leitura e manipulação de dados eficiente.

## ✨ Funcionalidades

* **Consolidação Rápida:** Junta o conteúdo de todos os arquivos `.xlsx` de um diretório em um único *DataFrame*.
* **Identificação de Fonte:** Adiciona automaticamente uma nova coluna chamada `Arquivo_Fonte` em cada linha, permitindo rastrear o documento original dos dados.
* **Saída Automática:** Cria um arquivo de saída com data e hora no nome para evitar sobrescrever execuções anteriores.

## ⚙️ Pré-requisitos

Para executar este script, você precisa ter o **Python 3.x** instalado e as seguintes bibliotecas:

1.  **pandas:** Para manipulação e análise de dados.
2.  **openpyxl:** Para leitura e escrita do formato `.xlsx` pelo Pandas.

Você pode instalar as dependências usando o `pip`:

```bash
pip install pandas openpyxl
