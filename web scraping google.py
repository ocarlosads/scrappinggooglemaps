from tkinter import scrolledtext
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import *
from selenium.webdriver.common.by import By
import csv
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
import requests
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.options import Options
from time import sleep
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
from google.auth import default
from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.common.by import By
import urllib3
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import scrolledtext
from tkinter import StringVar
from ttkthemes import ThemedStyle
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QVBoxLayout, QWidget, QLineEdit
from PyQt5.QtGui import QPalette, QColor
from qdarkstyle import load_stylesheet
from ttkbootstrap import Style
from ttkthemes import ThemedTk
itens = []




entry_nicho = None
entry_ddd = None


def iniciar_raspagem():
    nicho = entry_nicho.get()
    ddd = entry_ddd.get()
    

    # Configuração do WebDriver e opções do Chrome
    options = Options()
    options.add_argument('window-size=400,800')
    navegador = webdriver.Chrome(options=options)

    # URL inicial da pesquisa no Google
    url_base = f"https://www.google.com/search?sca_esv=580252979&rlz=1C1ONGR_pt-PTBR1083&tbs=lf:1,lf_ui:10&tbm=lcl&sxsrf=AM9HkKnXD9FykiaspKebeDZR4Y4254kLcA:1699403757248&q={nicho}&num=10&sa=X&ved=2ahUKEwjBj5-qlLOCAxVunpUCHXTPBykQjGp6BAgQEAE&biw=982&bih=707&dpr=1.25"

    # Lista para armazenar os resultados
    itens = []

    # Comece na primeira página
    pagina_atual = 1

    while True:
        # Navegar para a página atual
        url_pagina = url_base + f'&start={pagina_atual * 10}'
        navegador.get(url_pagina)
        sleep(2)  # Aguarde um momento para o carregamento da página

        # Obtenha a página da web
        pagina = navegador.page_source
        site = BeautifulSoup(pagina, 'html.parser')

        anuncios = site.findAll('div', attrs={'rllt__details'})

        for anuncio in anuncios:
            nome = anuncio.find('span', attrs={'OSrXXb'}).text
            notagoogle_element = anuncio.find('span', attrs={'yi40Hd YrbPuc'})

            if notagoogle_element is not None:
                notagoogle = notagoogle_element.text
            else:
                notagoogle = "sem nota"

            divs = anuncio.find_all('div')

            # ...

            for contato in divs:
                if ddd in contato.get_text():
                    div_text = contato.get_text()
                    posicao = div_text.find("·")

                    if posicao != -1:
                        texto_tratado = div_text[posicao + 1:]

                        # Remover parênteses e espaços do número de contato
                        texto_tratado = ''.join(c for c in texto_tratado if c.isdigit() or c == '+')
                    else:
                        texto_tratado = div_text

# ...


            # Verifique se o item já foi coletado
            texto_tratado = "55"+texto_tratado
            contato1 = f"https://wa.me/{texto_tratado}?text=Boa+noite%2C+tudo+bem%3F"
            item_unico = (nome, contato1, notagoogle)
            if item_unico not in itens:
                itens.append(item_unico)

        # Verifique a existência do botão "Próxima página"
        botao_proxima = site.find('span', style='display:block;margin-left:53px')

        if botao_proxima:
            # Se o botão "Próxima página" existir, clique nele e continue para a próxima página
            navegador.find_element(By.XPATH, '//*[@id="pnnext"]/span[2]').click()
            pagina_atual += 1
        else:
            # Se o botão "Próxima página" não existir, saia do loop
            break

    # Feche o navegador após a raspagem
    navegador.quit()

    dados = pd.DataFrame(itens, columns=['Empresa', 'Contato', 'Nota Google'])
    print(dados)

    
    gc = gspread.service_account(filename='C:\\Users\\carlo\\Desktop\\webscraping google\\webscraping-369322-3cdd189db15a.json')

    # Abra a planilha pelo nome ou URL
    planilha = gc.open("webscraping")

    # Selecione a guia na planilha onde deseja fazer upload dos dados
    nova_guia = planilha.add_worksheet(f"{nicho}", rows=1, cols=1)

    # Leve seu DataFrame para o Google Sheets
    df = pd.DataFrame(itens, columns=['Empresa', 'Contato', 'Nota Google'])
    dados = [df.columns.values.tolist()] + df.values.tolist()



    # Atualize a guia com os dados
    nova_guia.update(dados)

    print("Os dados foram coletados e enviados para o Google Sheets com sucesso!")

        # Restante do código para enviar dados ao Google Sheets

    messagebox.showinfo("Concluído", "Os dados foram coletados e enviados para o Google Sheets com sucesso!")
    
# Crie uma função para iniciar a interface gráfica
def criar_interface_grafica():
    global entry_nicho, entry_ddd  # Use as variáveis globais

    # Crie uma janela principal com o tema "superhero" e no modo escuro
    app = ThemedTk(theme="superhero", themebg=True)

    # Crie widgets com o estilo padrão do Tkinter
    label = ttk.Label(app, text="O que você deseja prospectar:")
    entry_nicho = ttk.Entry(app)
    label_ddd = ttk.Label(app, text="Qual o DDD da cidade:")
    entry_ddd = ttk.Entry(app)
    button_iniciar = ttk.Button(app, text="Iniciar", command=iniciar_raspagem)

    # Organize os widgets na janela
    label.pack(pady=10)
    entry_nicho.pack(pady=5)
    label_ddd.pack(pady=10)
    entry_ddd.pack(pady=5)
    button_iniciar.pack(pady=10)

    # Execute o loop principal da janela
    app.mainloop()

# Chame a função para criar a interface gráfica
criar_interface_grafica()































