import pyautogui
import pandas as pd
import os
import subprocess
import keyboard
from time import sleep
from pyautogui import ImageNotFoundException
import smtplib
from unidecode import unidecode
import pdfplumber
from dateutil.relativedelta import relativedelta
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.mime.image import MIMEImage
from datetime import datetime, timedelta
import calendar
import argparse
import json
import requests
from dotenv import load_dotenv
import cv2
import numpy as np
from PIL import ImageGrab
import ctypes 
import time
from datetime import datetime
import pytesseract
import easyocr



load_dotenv()

DB_LOGIN_FORTES = os.getenv('LOGIN_FORTES')
DB_SENHA_FORTES = os.getenv('SENHA_FORTES')
DB_CODIGO = os.getenv('CODIGO')

pasta_atual = os.path.dirname(os.path.abspath(__file__))

pasta_fotos = os.path.join(pasta_atual, "fotos")

arquivo1 = os.path.join(pasta_fotos, "gerar_guia_1.png")
arquivo2 = os.path.join(pasta_fotos, "mostrar_guia.png")
arquivo3 = os.path.join(pasta_fotos, "erro_geracao_guia.png")
arquivo4 = os.path.join(pasta_fotos, "sinal_gerada.png")
arquivo5 = os.path.join(pasta_fotos, "esperar_sistema_abrir5.png")
arquivo6 = os.path.join(pasta_fotos, "espera_outra_empresa22.png")
arquivo7 = os.path.join(pasta_fotos, "salvar_pdf.png")
arquivo8 = os.path.join(pasta_fotos, "botao_salvar.png")
arquivo9 = os.path.join(pasta_fotos, "aumentar_tela.png")
arquivo10 = os.path.join(pasta_fotos, "guia_ja_solicitada.png")
arquivo11 = os.path.join(pasta_fotos, "Contexto.png")

informacao1 = os.path.join(pasta_fotos, "guia_solicitada_sucesso.png")
informacao2 = os.path.join(pasta_fotos, "guia_ja_gerada.png")
informacao3 = os.path.join(pasta_fotos, "fgts_competemcia_mes.png")
informacao4 = os.path.join(pasta_fotos, "aguarde_guia_ja_solicitada.png")

advertencia1 = os.path.join(pasta_fotos, "lanc_advertencia_controle1.png")
advertencia2 = os.path.join(pasta_fotos, "botao_advertencia1_fechar2.png")
advertencia3 = os.path.join(pasta_fotos, "adv_atualizar_sistema3.png")
advertencia6 = os.path.join(pasta_fotos, "adv_3_dc_comunicado.png")
advertencia7 = os.path.join(pasta_fotos, "adv_7_dc_comunicado.png")
advertencia8 = os.path.join(pasta_fotos, "adv_8_dc_comunicado.png")

problema1 = os.path.join(pasta_fotos, "menu_bugado.png")
problema2 = os.path.join(pasta_fotos, "falha_att.png")

parametro1 = os.path.join(pasta_fotos, "parametro1.png")



dados = {
    "data": [
        {
            "cod_empresa":"0090",
            "nome_empresa": "CRECHE ESCOLA ESPACO VIDA",
            "local_arquivo": r"C:\Users\fisca\Music\Teste",
            "tipo_guia": "Mensal",
            "periodo": "06/2025",
            "tipo_debito": "FGTS e Consignado"
        },
        {
            "cod_empresa":"9112",
            "nome_empresa": "BABOLIM SERVICOS",
            "local_arquivo": r"C:\Users\fisca\Music\Teste",
            "tipo_guia": "Mensal",
            "periodo": "06/2025",
            "tipo_debito": "Somente FGTS"
        },
                {
            "cod_empresa":"0008",
            "nome_empresa": "AGF MEDICAL",
            "local_arquivo": r"C:\Users\fisca\Music\Teste",
            "tipo_guia": "Mensal",
            "periodo": "06/2025",
            "tipo_debito": "Somente Consignado"
        }
    ]
}


def formata_periodo(periodo):
    return periodo.replace('/', '-')



def verifica_menu(problema1, similaridade=0.8, limite=10):
    try:
        for tentativa in range(limite):
            posicao = localizar_elemento_cv(problema1, similaridade)
            if posicao:
                pyautogui.hotkey('alt')
                return
            else:
                sleep(1)
    except Exception as e:
        print(f"Erro ao verificar menu: {e}")
        
        




def localizar_elemento_cv(imagem_path, similaridade = 0.7):
    #carrega imagem alvo 
    imagem_alvo = cv2.imread(imagem_path)

    #captura Tela 
    screenshot = np.array(ImageGrab.grab())
    frame =  cv2.cvtColor(screenshot, cv2.COLOR_RGB2BGR)

    #faz o match 
    resultado = cv2.matchTemplate(frame, imagem_alvo, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(resultado)

    if max_val >= similaridade:
        h, w = imagem_alvo.shape[:2]
        centro_x = max_loc[0] + w //2
        centro_y = max_loc[1] + h//2
        return centro_x, centro_y
    else:
        return None
    

def mover_cursor_e_clicar(x, y):
    try:
        ctypes.windll.user32.SetCursorPos(x , y)
        time.sleep(0.2)
        ctypes.windll.user32.mouse_event(2,0,0,0,0)  # botão direito pressionado 
        ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)  # botão esquerdo solto
    except Exception as e:
        print(f"Erro ao mover o cursor e clicar: {e}")
        raise e




def esperar_cv(imagem_path, similaridade=0.7, limite=600):
    try:
        cont = 0
        while cont < limite:
            posicao = localizar_elemento_cv(imagem_path, similaridade)
            if posicao:
                return posicao
            cont += 1
            time.sleep(1)
        print(f"Imagem {imagem_path} não encontrada após {limite} tentativas.")
        return None
    except Exception as e:
        print(f"Erro ao esperar pela imagem {imagem_path}: {e}")
        raise ImageNotFoundException(f"Imagem {imagem_path} não encontrada após {limite} tentativas.") from e



def fechar_fortes():
    try:
        # Comando para finalizar o processo Fortes
        os.system("taskkill /f /im AC.exe")
        print("O programa Fortes foi fechado com sucesso.")
    except Exception as e:
        print(f"Erro ao tentar fechar o programa Fortes: {e}")


def advertencia_1(advertencia1, advertencia2, similaridade=0.8):
        pos1 = esperar_cv(advertencia1, similaridade, limite=10)
        if pos1:
            sleep(3)
            pos2 = esperar_cv(advertencia2,similaridade, limite=10)
            if pos2:
                mover_cursor_e_clicar(*pos2)
                sleep(3)
                pyautogui.hotkey('alt')
        else:
            print(f"Imagem {advertencia1} não encontrada")
        


def advertencia_3(advertencia3, similaridade = 0.8):
        pos1 = esperar_cv(advertencia3, similaridade, limite=10)
        if pos1:
            sleep(3)
            pyautogui.press('tab')
            sleep(3)
            pyautogui.press('enter')
            sleep(3)
        else:
            print(f"Imagem {advertencia3} não encontrada")
        


def advertencia_6(advertencia8, similaridade=0.8):
    imagens = [advertencia8]
    for imagem in imagens:
        pos = esperar_cv(imagem, similaridade, limite=10)
        if pos: 
            sleep(3)
            pyautogui.press('enter')
            sleep(3)
    else: 
        print(f"Imagem {imagem} não encontrada")


def capturar_icone_inicial():
    bbox = (1183, 41, 1265, 59)  
    screenshot = np.array(ImageGrab.grab(bbox=bbox))
    screenshot_bgr = cv2.cvtColor(screenshot, cv2.COLOR_RGB2BGR)
    caminho = "primeiro_print.png"
    cv2.imwrite(caminho, screenshot_bgr)
    return screenshot_bgr

def capturar_icone_final():
    bbox = (1183, 41, 1265, 59)  
    screenshot = np.array(ImageGrab.grab(bbox=bbox))
    screenshot_bgr = cv2.cvtColor(screenshot, cv2.COLOR_RGB2BGR)
    caminho = "imagem_loop.png"
    cv2.imwrite(caminho, screenshot_bgr)
    return screenshot_bgr    

def espera_componente_mudar(problema2, similaridade=0.8, intervalo=5):
    try:
        primeira_linha = capturar_icone_inicial()

        while True:
            sleep(intervalo)
            pyautogui.press('f5')
            sleep(2)  # Aguarda a tela recarregar

            # Se aparecer o problema2, aperta enter e continua esperando
            pos_problema2 = localizar_elemento_cv(problema2, similaridade)
            if pos_problema2:
                pyautogui.press('enter')
                sleep(2)
                

            icone_final = capturar_icone_final()
            diferenca = cv2.absdiff(primeira_linha, icone_final)
            soma_diferenca = np.sum(diferenca)

            if soma_diferenca > 1000:  

                return True 
    except Exception as e:
        print(f"Erro ao esperar componente mudar: {e}")
        return False

def baixar_pdf(arquivo2, similaridade=0.8, limite=30, intervalo=8):
    try:
        for tentativa in range(limite):
            sleep(intervalo)             
            pos_pdf_guia = localizar_elemento_cv(arquivo2, similaridade)
            if pos_pdf_guia:
                mover_cursor_e_clicar(*pos_pdf_guia)
                sleep(2)
                return True
            # Verifica se apareceu a imagem de erro
            pos_erro = localizar_elemento_cv(arquivo3, similaridade)
            if pos_erro:
                print("Guia não gerada, pulando para a próxima empresa")
                return False
        print("Pdf não aberto após mudança na primeira linha")
        return False
    except Exception as e:
        print(f"Erro ao baixar PDF: {e}")
        return False


def salvar_arquivo(arquivo7, arquivo8, local_arquivo, similaridade=0.8, limite=30):
    try:
        for tentativa in range(limite):
            pos_salvar_pdf = localizar_elemento_cv(arquivo7, similaridade)
            if pos_salvar_pdf:
                mover_cursor_e_clicar(*pos_salvar_pdf)
                sleep(5)
                novo_periodo = formata_periodo(periodo)
                pyautogui.write(f"Guia_FGTS_{novo_periodo}_{nome_empresa}")
                sleep(5)
                pyautogui.hotkey('ctrl', 'l')
                sleep(3)
                pyautogui.write(local_arquivo)
                sleep(3)
                pyautogui.press('enter')
                sleep(3)
                for tentativa_botao in range(limite):
                    pos_salvar_botao = localizar_elemento_cv(arquivo8, similaridade)
                    if pos_salvar_botao:
                        mover_cursor_e_clicar(*pos_salvar_botao)
                        print(f"Arquivo salvo na pasta {local_arquivo}")
                        return True
                    else:
                        print("Arquivo não salvo")
                return False
                sleep(1)
        return False
    except Exception as e:
        print(f"Erro ao salvar arquivo: {e}")
        return False
    

def fullscreen(arquivo9, similaridade=0.8, limite=10):
    try:
        for tentativa in range(limite):
            pos_tela_cheia = localizar_elemento_cv(arquivo9, similaridade)
            if pos_tela_cheia:
                mover_cursor_e_clicar(*pos_tela_cheia)
                print("fullscreen")
                return
            else:
                print("Tentando fullscreen novamente...")
                sleep(1)
        print("não ta fullscreen")
    except Exception as e:
        print(f"Erro ao tentar fullscreen: {e}")
    

def mudando_contexto(arquivo11, similaridade=0.8, limite=30):
    try:
        for tentativa in range(limite):
            pos_contexto = localizar_elemento_cv(arquivo11, similaridade)
            if pos_contexto:
                mover_cursor_e_clicar(*pos_contexto)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('tab')
                pyautogui.write(periodo_preencher)
                sleep(2)
                pyautogui.hotkey('tab')
                pyautogui.write(periodo_preencher)
                sleep(2)
                pyautogui.hotkey('f9')
                sleep(5)
                return
                sleep(2)
        print("Contexto não encontrado")
    except Exception as e:
        print(f"Erro ao mudar contexto: {e}")
        

def pegar_linha_mais_proxima_primeira():

    screenshot = np.array(ImageGrab.grab(bbox=(14, 43, 1274, 61)))
    imagem = cv2.cvtColor(screenshot, cv2.COLOR_RGB2BGR)
    
    caminho = "linha_inicial.png"
    cv2.imwrite(caminho, imagem)

    # Usa EasyOCR para ler o texto diretamente do array da imagem
    reader = easyocr.Reader(['pt'], gpu=False)
    results = reader.readtext(imagem, detail=0)
    texto = " ".join(results).strip()


    if not texto:
        return None

    print("✅ Primeira linha capturada:")
    print(texto)
    return texto



exe_path = r"C:\Fortes\AC\AC.exe"
start_in = r"\\CONTABSERVER\Fortes\AC"

try : 
    # abrir sistema fortes ###PRIMEIRA PARTE####
    subprocess.Popen([exe_path], cwd=start_in)
    pos = esperar_cv(arquivo5)
    #if pos:
    #    mover_cursor_e_clicar(*pos)
    sleep(2)
    # preencher campo usuário
    pyautogui.write(DB_LOGIN_FORTES)
    sleep(2)
    # apertar botao tab
    pyautogui.press('tab')
    sleep(2)
    # preencher campo senha
    pyautogui.write(DB_SENHA_FORTES)
    sleep(2)
    # apertar botao tab
    pyautogui.press('tab')
    sleep(2)
    # apertar botao page down
    pyautogui.press('pagedown')
    sleep(2)
    # apertar botao 3
    pyautogui.hotkey('2')
    sleep(2)
    # apertar botao enter
    pyautogui.press('enter')
    sleep(2)
    # preencher campo codigo empresa so para abrir sistema
    pyautogui.write(DB_CODIGO)
    sleep(2)
    # apertar botao tab
    pyautogui.press('tab')
    sleep(2)
    # apertar botao f9
    pyautogui.press('f9')
    sleep(10)
    advertencia_6(advertencia8) 
    sleep(2)
    advertencia_3(advertencia3)
    sleep(2)
    advertencia_1(advertencia1, advertencia2)
    sleep(2)
    # ativar a tela do sistema fortes novamente
    pyautogui.hotkey('alt')
    #x, y = pyautogui.locateCenterOnScreen(arquivo1, confidence=0.7)
    #pyautogui.click(x, y)
    sleep(2)

    for dado in dados['data']:
        cod_empresa = dado['cod_empresa']
        nome_empresa = dado['nome_empresa']
        local_arquivo = dado['local_arquivo']
        tipo_guia = dado['tipo_guia']
        periodo = dado['periodo']
        tipo_debito = dado['tipo_debito']




        cam_geracao_guia = f'Movimentos / FGTS Digital / Painel FGTS Digital'

        sleep(3)
        pyautogui.hotkey('ctrl', 'e')
        pos_outra_empres = esperar_cv(arquivo6)

        pyautogui.write(cod_empresa)
        sleep(2)
        pyautogui.press('tab')
        sleep(2)
        pyautogui.press('f9')
        sleep(10)
        pyautogui.hotkey('alt')
        sleep(5)
        verifica_menu(problema1, limite=10)
        sleep(5)
        pyautogui.write(cam_geracao_guia)
        sleep(2)
        pyautogui.hotkey('enter')

        sleep(8)
        fullscreen(arquivo9)
        sleep(5)
        
        pos_gerar_guia = esperar_cv(arquivo1)
        if pos_gerar_guia: 
            mover_cursor_e_clicar(*pos_gerar_guia)
            print("Botão Clicado com sucesso")
        else:
            print("Botão não encontrado na tela")


        if tipo_guia == "Anual":
            periodo_preencher = periodo[-4:]
        else:
            periodo_preencher = periodo

        sleep(5)
        pyautogui.write(tipo_guia)
        sleep(2)
        pyautogui.hotkey('tab')
        pyautogui.write(periodo)
        sleep(2)
        pyautogui.hotkey('tab')
        if tipo_debito == "FGTS e Consignado":
            pyautogui.press('pagedown')
            sleep(2)
            pyautogui.hotkey('enter')
        elif tipo_debito == "Somente FGTS":
            pyautogui.press('pagedown')
            sleep(2)
            pyautogui.hotkey('down')
            sleep(2)
            pyautogui.hotkey('enter')
        elif tipo_debito == "Somente Consignado":
            pyautogui.press('pagedown')
            sleep(2)
            pyautogui.hotkey('down')
            sleep(2)
            pyautogui.hotkey('down')
            sleep(2)
            pyautogui.hotkey('enter')
        else: 
            print("Tipo de débito não reconhecido, usando padrão 'FGTS e Consignado'")
        sleep(5)
        pyautogui.hotkey('f9')

        #mensagem de aguarde 
        try:
            pos_aguarde_guia = esperar_cv(informacao4, limite=10)
            if pos_aguarde_guia:
                print("guia ja gerada, por favor aguarde")
                sleep(2)
                pyautogui.hotkey('tab')
                sleep(2)
                pyautogui.hotkey('enter')
                sleep(2)
        except Exception as e:
            print("Guia não gerada, continuando...")    

        #mensagem de guia ja gerada
        try:
            pos_guia_ja_gerada = esperar_cv(informacao2, limite=10)
            if pos_guia_ja_gerada:
                pyautogui.hotkey('tab')
                sleep(2)
                pyautogui.hotkey('enter')
                sleep(2)
        except Exception as e:
            print("Guia não gerada, continuando...")

        #mensagem de guia ja solicitada
        try:
            pos_guia_ja_solicitada = esperar_cv(arquivo10, limite=10)
            if pos_guia_ja_solicitada:
                mover_cursor_e_clicar(*pos_guia_ja_solicitada)
                print("Guia já solicitada")
                sleep(2)
                pyautogui.hotkey('enter')
                sleep(2)
        except Exception as e:
            print("Guia não solicitada, continuando...")

        
        #guia com copetencia errada 
        try:
            pos_competencia_errada = esperar_cv(informacao3, limite=10)
            if pos_competencia_errada:
                pyautogui.hotkey('enter')
                print("Para gerar guia no FGTS Digital, encerre antes a transmissão do eSocial na competência {periodo}")
                sleep(2)
                pyautogui.hotkey('esc')
                sleep(2)
                pyautogui.hotkey('esc')
                sleep(2)
                continue
        except Exception as e:
            print("Competência correta, continuando...")


       #mensagem de guia gerada         
        
        pos_guia_gerada = esperar_cv(informacao1, limite=10)
        if pos_guia_gerada:
            pyautogui.hotkey('enter')
            print("Guia gerada com sucesso!")
            sleep(2)
        else: 
            print("Guia não gerada")

        #mudar periodo do contexto para a data que colocaram no servidor 
        mudando_contexto(arquivo11, limite=10)
        sleep(5)

        primeira_linha = pegar_linha_mais_proxima_primeira()
        sleep(5)

        checando = espera_componente_mudar(problema2)
        if checando is False: 
            pyautogui.hotkey('esc')
            sleep(2)
        sleep(5)

        pdf_dowload = baixar_pdf(arquivo2)
        if pdf_dowload is False:
            pyautogui.hotkey('esc')
            sleep(2)
            continue
        sleep(5)

        salvar_arquivo(arquivo7, arquivo8, local_arquivo)
        sleep(5)

        pyautogui.hotkey('esc')
        sleep(2)
        pos_parametro = localizar_elemento_cv(parametro1)
        if pos_parametro:
            mover_cursor_e_clicar(*pos_parametro)
            sleep(2)
            pyautogui.hotkey('esc')
            sleep(2)

except Exception as e:
    print(f"Erro no processamento principal: {e}")
    fechar_fortes()

finally:
    # Fecha o Fortes após o término ou em caso de erro
    fechar_fortes()