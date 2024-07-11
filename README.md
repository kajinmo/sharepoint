# SharePoint Access via Python

Este projeto fornece um script Python para acessar e interagir com o SharePoint usando a biblioteca `office365`. O script permite autenticação, operações com arquivos (como download, upload e listagem de arquivos) e recuperação de itens de listas do SharePoint.

## Funcionalidades

- Autenticação no SharePoint
- Listar arquivos em uma pasta do SharePoint
- Fazer download de arquivos
- Fazer upload de arquivos (incluindo upload em partes)
- Recuperar o arquivo mais recente com base na data de modificação
- Obter propriedades de arquivos de uma pasta
- Atualizar arquivos de fundo com os nomes de arquivos XML mais recentes
- Fazer download e ler arquivos Excel

## Requisitos

- Python 3.6+
- `office365-runtime`
- `pandas`
- `python-dotenv`

## Instalação

1. Clone o repositório:

    ```bash
    git clone https://github.com/kajinmo/sharepoint-access-python.git
    cd sharepoint-access-python
    ```


2. Instale os pacotes Python necessários:
    ```bash
    pip install -r requirements.txt
    ```

3. Crie um arquivo .env no diretório raiz e adicione suas credenciais e informações do site SharePoint:
    ```bash
    SHAREPOINT_EMAIL=seu-email@example.com
    SHAREPOINT_PASSWORD=sua-senha
    SHAREPOINT_URL_SITE=https://sua-url-do-sharepoint
    SHAREPOINT_SITE_NAME=nome-do-seu-site-sharepoint
    SHAREPOINT_DOC_LIBRARY=nome-da-sua-biblioteca-de-documentos
    ```

4. Crie um arquivo `.gitignore` no diretório raiz para garantir que suas credenciais não sejam incluídas no controle de versão:

    ```plaintext
    .env
    __pycache__/
    *.pyc
    *.pyo
    .DS_Store
    ```

## Uso

### Autenticação

O script usa `ClientContext` da biblioteca `office365` para autenticar no SharePoint usando as credenciais fornecidas.

### Listar Arquivos

Para listar todos os arquivos em uma pasta específica do SharePoint:

  ```python
  sharepoint = SharePoint()
  files = sharepoint.get_files_list('nome-da-sua-pasta')
  for file in files:
      print(file.name)
  ```

### Fazer Download de um Arquivo

Para fazer download de um arquivo específico do SharePoint:

  ```python
  file_content = sharepoint.download_file('nome-do-seu-arquivo', 'nome-da-sua-pasta')
  with open('nome-do-arquivo-baixado', 'wb') as f:
      f.write(file_content)
  ```

### Fazer Upload de um Arquivo

Para fazer upload de um arquivo para o SharePoint:

  ```python
  with open('nome-do-arquivo-local', 'rb') as f:
      content = f.read()
  sharepoint.upload_file('nome-do-seu-arquivo', 'nome-da-sua-pasta', content)
  ```

### Fazer Upload de um Arquivo em Partes

Para fazer upload de um arquivo grande em partes:

  ```python
  sharepoint.upload_file_in_chunks('caminho-para-arquivo-grande', 'nome-da-sua-pasta', chunk_size=1024*1024*5)
  ```

### Recuperar Arquivo Mais Recente

Para recuperar o arquivo mais recente de uma pasta específica do SharePoint:

  ```python
  latest_file_name, latest_file_content = sharepoint.download_latest_file('nome-da-sua-pasta')
  ```

### Recuperar Itens de uma Lista

Para recuperar itens de uma lista do SharePoint:

  ```python
  items = sharepoint.get_list('nome-da-sua-lista')
  for item in items:
      print(item.properties)
  ```

### Fazer Download e Ler Arquivo Excel

Para fazer download e ler um arquivo Excel do SharePoint:

  ```python
  df = sharepoint.download_and_read_excel('nome-do-seu-arquivo-excel', 'nome-da-sua-pasta')
  print(df.head())
  ```

## Contribuição
Contribuições são bem-vindas! Por favor, abra uma issue ou envie um pull request para quaisquer melhorias ou correções de bugs.
