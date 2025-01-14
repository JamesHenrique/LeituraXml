# Leitura XML

Este projeto é uma aplicação em Python com interface gráfica (GUI) desenvolvida para processar e formatar arquivos XML relacionados a diferentes categorias de eventos trabalhistas e fiscais, como S-2299, S-1200, entre outros. Ele automatiza a busca, análise e criação de planilhas formatadas com os dados extraídos.

## Funcionalidades

- Identifica e processa arquivos XML específicos dentro de uma pasta selecionada pelo usuário.
- Cria planilhas Excel para os seguintes eventos:
  - **S-2299**: Rescisão.
  - **S-1200**: Recibos.
  - **S-2200**: Admissão.
  - **S-2205**: Inclusão de Dependente.
  - **S-1010**: Verbas.
  - **S-2206**: Alteração de Cargo.
  - **S-5011**: Informações detalhadas.
- Apresenta mensagens de status em uma interface intuitiva, informando o progresso e os resultados.
- Interface gráfica em modo escuro, desenvolvida com a biblioteca `customtkinter`.

## Tecnologias Utilizadas

- **Linguagem**: Python.
- **GUI**: `customtkinter`.
- **Manipulação de Arquivos XML**: 
  - `xmltodict` para conversão de XML para dicionários.
  - `xml.dom.minidom` para manipulação de estrutura XML.
- **Manipulação de Dados**: `pandas` para criação e formatação de planilhas.
- **Manipulação de Arquivos**: `os` para gerenciar diretórios e arquivos.
- **Logs**: Biblioteca padrão `logging`.

## Pré-requisitos

Certifique-se de que seu ambiente possui as seguintes dependências instaladas:

- Python 3.8 ou superior.
- Bibliotecas:
  - `customtkinter`
  - `xmltodict`
  - `pandas`
  - `openpyxl` (necessária para manipulação de arquivos Excel).
  - Outras dependências mencionadas nos módulos personalizados.

## Instalação

1. Clone este repositório para o seu ambiente local:
   ```bash
   git clone https://github.com/seu-usuario/leitura-xml.git
   cd leitura-xml


2. Instale as dependências necessárias:
   pip install customtkinter xmltodict pandas openpyxl

1. **Execute o arquivo principal**:
 ```bash
 python app.py
 ```
2. **Na interface gráfica**:
   - Clique no botão **"ESCOLHER A PASTA"**.
   - Selecione a pasta que contém os arquivos XML desejados.
   - Aguarde enquanto o programa processa os arquivos. O progresso será exibido no painel de mensagen3. **Resultado**:
   - Após o processamento, as planilhas geradas serão salvas na mesma pasta onde os XMLs foram enco---
  ## Logs e Mensagens
  O programa exibe mensagens em tempo real na interface, informando:
  - Arquivos encontrados.
  - Planilhas criadas.
  - Erros ou arquivos não encontrados.
  ---


   
