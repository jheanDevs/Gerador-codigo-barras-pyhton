# Gerador de Códigos de Barras

Este projeto é uma aplicação desktop simples desenvolvida em **Python** utilizando a biblioteca **Flet** para gerar códigos de barras a partir de uma planilha Excel. A aplicação exibe o progresso e o status de cada código de barras gerado, indicando se o código já existe ou se foi criado com sucesso.

## Funcionalidades

- **Upload de Planilha Excel**: O usuário pode selecionar uma planilha Excel com os dados dos funcionários.
- **Geração de Códigos de Barras**: A aplicação gera automaticamente códigos de barras para cada funcionário baseado em sua matrícula (CHAPA).
- **Exibição de Log**: A aplicação exibe em tempo real o progresso do processo de geração de códigos de barras e registra o status de cada funcionário.
- **Barra de Progresso**: Mostra visualmente o progresso da geração dos códigos.
- **Controle de Duplicatas**: Caso um código de barras já exista, ele não será gerado novamente.
- **Alerta de Conclusão**: Um pop-up é exibido ao final do processo informando quantos novos códigos foram gerados.

## Tecnologias Utilizadas

- **Python**: Linguagem principal utilizada no projeto.
- **Flet**: Framework para construção da interface gráfica.
- **Pandas**: Utilizado para manipulação da planilha Excel.
- **Python Barcode**: Biblioteca usada para gerar os códigos de barras no formato Code128.
- **OpenPyXL**: Para leitura de arquivos Excel.

## Requisitos

- **Python 3.7+**
- **Bibliotecas Python**:
  - `flet`
  - `pandas`
  - `openpyxl`
  - `python-barcode`
  
## Instalação

1. **Clone o repositório**:
   ```bash
   git clone https://github.com/usuario/gerador-codigos-barras.git
   cd gerador-codigos-barras
