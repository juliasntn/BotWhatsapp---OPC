
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import os 
import pyautogui, sys
import time
import schedule
import logging
import datetime as dt
from opcua import Client as opcclient



    # Configurar o nível de registro
logging.basicConfig(filename='BOT.log',
                        filemode='a',
                        format='%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s',
                        encoding='utf-8',
                        datefmt='%Y-%m-%d %H:%M:%S',
                        level=logging.DEBUG)

    # Criar um logger
logger = logging.getLogger("BotWhatsapp")

    # Definir um formato para as mensagens de registro
formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")

    # Criar um manipulador de console
console_handler = logging.StreamHandler()
console_handler.setFormatter(formatter)

    # Adicionar o manipulador ao logger
logger.addHandler(console_handler)

    # Exemplos de mensagens de registro
'''
    logger.debug("Esta é uma mensagem de depuração")
    logger.info("Esta é uma mensagem de informação")
    logger.warning("Esta é uma mensagem de aviso")
    logger.error("Esta é uma mensagem de erro")
    logger.critical("Esta é uma mensagem crítica")
'''

logger.info("Iniciando tarefa....")
opc_server_url = "opc.tcp://SERVIDOR"
def tarefa_programada():
    logger.info("Tarefa iniciada, código rodando...")
        #opc servidor
    

        # Lista de tags que você deseja ler
    tags_to_read = [
            {"name":"NOME DA TAG", "tag": "ns=2;s=TAG"},
            {"name":"NOME DA TAG", "tag": "ns=2;s=TAG"},
            {"name":"NOME DA TAG","tag":"ns=2;s=TAG"},
            {"name":"NOME DA TAG","tag":"ns=2;s=TAG"}
        ]

    tag_ranges  = {
            "ns=2;s=TAG":("VALOR MIN","VALOR MAX"),  
            "ns=2;s=TAG":("VALOR MIN","VALOR MAX"),
            "ns=2;s=TAG":("VALOR MIN","VALOR MAX"),
            "ns=2;s=TAG":("VALOR MIN","VALOR MAX")
        }

        # Lista para armazenar as informações de cada tag
    tag_info = []

    try:
            logger.info("Conectando ao servidor OPC UA")
            # Conecte-se ao servidor OPC UA
            client_opc = opcclient("opc.tcp://SERVIDOR")
            client_opc.connect()

            # Leia os valores das tags OPC UA e verifique se estão dentro dos intervalos especificados
            for tag_data in tags_to_read:
                tag_name = tag_data["name"]
                tag_tag = tag_data["tag"]
                try:
                    node = client_opc.get_node(tag_tag)
                    value = round(node.get_value(), 2)  # Arredonda o valor para duas casas decimais
                    if tag_tag in tag_ranges:
                        min_value, max_value = tag_ranges[tag_tag]
                        if min_value <= value <= max_value:
                            status = "✅"
                        else:
                            status = "⚠️"
                        tag_info.append(f"{tag_name}: {value} - {status}")
                except Exception as e:
                    print(f"Erro ao ler a tag {tag_tag}: {e}")
                    logger.error(f"Erro ao ler as tags: {e}")

            # Compile todas as informações
            message = "\n".join(tag_info)



            # Desconecte-se do servidor OPC UA
            client_opc.disconnect()
            logger.info("Desconectado do servidor OPC UA")

    except Exception as ex:
            print(f"Erro ao conectar ao servidor OPC UA: {ex}")
            logging.error(f"Erro ao conectar ao servidor OPC UA: {ex}")


    webbrowser.open('https://web.whatsapp.com/')
    sleep(30)

    # Ler planilha e guardar informações sobre nome, telefone e data de vencimento
    workbook = openpyxl.load_workbook('bot whatsapp/numeros.xlsx')
    pagina_clientes = workbook['Sheet1']

    for linha in pagina_clientes.iter_rows(min_row=2):
        # nome, telefone, vencimento
        nome = linha[0].value
        telefone = linha[1].value
        
        mensagem = f'Olá {nome} isso é um teste'

        # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
        # com base nos dados da planilha
        try:
            link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(message)}'
            webbrowser.open(link_mensagem_whatsapp)
            sleep(10)
            seta = pyautogui.locateCenterOnScreen('bot whatsapp/seta.png')
            sleep(2)
            pyautogui.click(seta[0],seta[1])
            sleep(2)
            pyautogui.hotkey('ctrl','w')
            sleep(2)
        except:
            print(f'Não foi possível enviar mensagem para {nome}')
            with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
                arquivo.write(f'{nome},{telefone}{os.linesep}')
    
schedule.every().hours.do(tarefa_programada)

while True:
    schedule.run_pending()
    time.sleep(1)
