### README

# Bot de WhatsApp Automatizado

Este projeto é um bot de WhatsApp automatizado que lê informações de um servidor OPC UA e envia mensagens para contatos especificados em uma planilha Excel. As mensagens são enviadas via WhatsApp Web.

## Funcionalidades

- Conexão a um servidor OPC UA para leitura de tags específicas.
- Verificação de valores de tags contra intervalos predefinidos.
- Envio automático de mensagens de WhatsApp para contatos listados em uma planilha Excel.
- Registro de logs detalhados para monitoramento de atividades e erros.

## Requisitos

- Python 3.8 ou superior
- Bibliotecas Python:
  - `openpyxl`
  - `urllib`
  - `webbrowser`
  - `time`
  - `os`
  - `pyautogui`
  - `schedule`
  - `logging`
  - `datetime`
  - `opcua`

## Instalação

1. Clone o repositório:

    ```bash
    git clone https://github.com/seu-usuario/seu-repositorio.git
    ```

2. Navegue até o diretório do projeto:

    ```bash
    cd seu-repositorio
    ```

3. Instale as dependências:

    ```bash
    pip install openpyxl pyautogui schedule opcua
    ```

## Configuração

1. Atualize o URL do servidor OPC UA no código:

    ```python
    opc_server_url = "opc.tcp://SEU_SERVIDOR_OPC"
    ```

2. Atualize a lista de tags e seus intervalos no código:

    ```python
    tags_to_read = [
        {"name": "NOME_DA_TAG", "tag": "ns=2;s=TAG"},
        ...
    ]

    tag_ranges = {
        "ns=2;s=TAG": ("VALOR_MIN", "VALOR_MAX"),
        ...
    }
    ```

3. Prepare a planilha `numeros.xlsx` com as colunas:
    - Nome
    - Telefone

4. Salve um ícone de seta (`seta.png`) para auxiliar o PyAutoGUI no envio das mensagens via WhatsApp Web.

## Uso

Execute o script:

```bash
python bot.py
```

O bot irá:

1. Conectar-se ao servidor OPC UA e ler as tags especificadas.
2. Verificar os valores das tags contra os intervalos definidos.
3. Compilar as informações e enviar mensagens de WhatsApp para os contatos listados na planilha `numeros.xlsx`.

## Logs

Os logs são salvos no arquivo `BOT.log` e exibidos no console para facilitar a depuração e o monitoramento das atividades do bot.

## Contribuição

Sinta-se à vontade para fazer um fork do projeto, criar uma nova branch e enviar pull requests. 

## Contato
linkedin: https://www.linkedin.com/in/julia-santana-040a12180/
