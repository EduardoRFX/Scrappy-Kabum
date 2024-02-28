from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import xlsxwriter
import os
from colorama import init, Fore 
init()

def navegador_config():
    opcoes = webdriver.ChromeOptions()
    opcoes.add_experimental_option("detach", True)
    opcoes.add_experimental_option('excludeSwitches', ['enable-logging'])

    navegador = webdriver.Chrome(options=opcoes)
    return navegador

def excel_config():
    caminho = '<your.path.example>'
    planilha = xlsxwriter.Workbook(filename=caminho)

    return planilha


class Scrappy():

    def __init__(self):
        self.nomes_cadeiras = []
        self.valores_cadeiras = []

    def iniciar(self):
        self.raspagem_cadeiras()
        self.envio_excel()

    def raspagem_cadeiras(self):
        navegador = navegador_config()
        navegador.set_window_size(1360,760)
        navegador.get('https://www.kabum.com.br/espaco-gamer/cadeiras-gamer?page_number=1&page_size=20&facet_filters=&sort=most_searched')
        # wait = WebDriverWait(navegador, 10)

        posi = 0
        for p in range(3, 7):
            item = 1
            for i in range(1,11):
                nome_cadeira = navegador.find_elements(By.XPATH, f'//*[@id="listing"]/div[3]/div/div/div[2]/div[1]/main/div[{item}]/a/div/button/div/h2/span')[0].text
                self.nomes_cadeiras.append(nome_cadeira)

                # sleep(1)

                valor_cadeira = navegador.find_elements(By.XPATH, f'//*[@id="listing"]/div[3]/div/div/div[2]/div[1]/main/div[{item}]/a/div/div/span[2]')[0].text
                self.valores_cadeiras.append(valor_cadeira)

                print(f'Nome da Cadeira: {self.nomes_cadeiras[posi]}')
                print(f'Valor da Cadeira: {self.valores_cadeiras[posi]}\n')
                item += 1
                posi += 1

                # sleep(1)
            
            botao_proxima_pagina = navegador.find_element(By.XPATH, f'//*[@id="listingPagination"]/ul/li[{p}]/a')
            botao_proxima_pagina.click()

            print(Fore.GREEN + '\nPróxima Página.............\n')
            print(Fore.RESET)
            sleep(4)
        
    def envio_excel(self):
        planilha = excel_config()
        sheet = planilha.add_worksheet('Scrappy')

        formato_titulo = planilha.add_format({
            'bg_color':'gray',
            'font_color': 'white',
            'align': 'center',
            'bold': True,
        })

        sheet.set_column('A:A',127)
        sheet.set_column('B:B', 11)

        sheet.write('A1', 'NOME', formato_titulo)
        sheet.write('B1', 'VALOR', formato_titulo)
        
        sheet.write_column('A2', self.nomes_cadeiras)
        sheet.write_column('B2', self.valores_cadeiras)

        planilha.close()
        os.startfile('<path.example>')

scrappy = Scrappy()
scrappy.iniciar()
