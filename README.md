# Automação SAP GUI com Python

Este projeto foi desenvolvido para automatizar a extração de relatórios do SAP GUI utilizando Python e a biblioteca `pywin32`. A automação contempla a execução de diversas transações SAP (MB51, MB25, ZMM304 e MC.9), salvando os arquivos gerados no Excel e posteriormente movendo-os para diretórios de rede específicos.

## 🧾 Funcionalidades

- Conecta automaticamente ao SAP GUI.
- Executa transações:
  - **MB51** – Movimentações de material.
  - **MB25** – Requisições de saída.
  - **ZMM304** – Relatório de empréstimos.
  - **MC.9** – Análise de consumo parcial.
- Exporta os dados diretamente para planilhas Excel.
- Fecha os arquivos Excel automaticamente após a exportação.
- Move os arquivos gerados para pastas específicas da rede.

## ⚙️ Pré-requisitos

Antes de rodar o script, é necessário garantir que:

- O SAP GUI está instalado e configurado corretamente com acesso via scripting.
- O scripting do SAP GUI está habilitado (servidor e cliente).
- A conta de usuário tem acesso às transações mencionadas.
- As bibliotecas abaixo estejam instaladas:

```bash
pip install pywin32 psutil
