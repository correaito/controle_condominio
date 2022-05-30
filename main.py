import tkinter as tk
from tkinter import ANCHOR, Entry, ttk, Label
from tkinter.tix import ExFileSelectBox, Tk
from tkinter import messagebox
import time
from datetime import datetime
from numpy import column_stack
from setuptools import Command
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from subprocess import CREATE_NO_WINDOW
import codecs
import pandas as pd
import math
from operator import neg
import sys

# aqui o pandas irá importar os dados da planilha
def read_excel():
    condominio_df = pd.read_excel('Calculos.xlsx')
    novo_condominio_df = condominio_df.set_index('ID')
    return novo_condominio_df

# mensagem de confirmação ao fechar o programa
def ao_fechar():
    janela_sair = tk.Tk()
    janela_sair.title("Antes de Sair")

    texto_label = Label(janela_sair, text="Deseja salvar o Histórico? Clique em Sim ou Não", bg='black', fg='white', relief='solid')
    texto_label.grid(row=0, column=0, padx=10, pady=10,sticky='nsew', columnspan=2)

    botao_continuar = tk.Button(
        janela_sair, text="Sim", command=grava_historico, bg='green', fg='white', relief='solid')
    botao_continuar.grid(row=2, column=0, padx=10, pady=10, sticky='nsew')

    botao_sair = tk.Button(
        janela_sair, text="Não", command=finaliza_programa)
    botao_sair.grid(row=2, column=1, padx=10, pady=10, sticky='nsew')
    janela_sair.mainloop()
    
def finaliza_programa():
    sys.exit()

# essa funcao vai zerar todas as despesas e saldos da planilha
def limpar_linhas():
    novo_condominio_df = read_excel()

    for coluna in range(10):
        coluna += 1
        for linha in range(7):
            novo_condominio_df.iloc[linha, coluna] = 0
    novo_condominio_df.to_excel('Calculos.xlsx')

    janela_finalizacao_limpeza = tk.Tk()
    janela_finalizacao_limpeza.withdraw()
    tk.messagebox.showinfo("App Edificio Guanabara",
                           "Ok, planilha limpa com sucesso!")

# define o mês de competencia automaticamente
def define_competencia():
    novo_condominio_df = read_excel()

    data_atual = datetime.now()
    # aqui faremos uma conversao do mes (numeral) para mes (por extenso)
    mes_ext = {1: 'jan', 2: 'fev', 3: 'mar', 4: 'abr', 5: 'mai', 6: 'jun',
               7: 'jul', 8: 'ago', 9: 'set', 10: 'out', 11: 'nov', 12: 'dez'}
    mes_atual = data_atual.strftime('%m/%Y')
    mes, ano = mes_atual.split("/")
    mes = '{}/{}'.format(mes_ext[int(mes)], ano)

    for linha in range(6):
        novo_condominio_df.iloc[linha, 10] = mes
    novo_condominio_df.to_excel('Calculos.xlsx')

    janela_finalizacao_compt = tk.Tk()
    janela_finalizacao_compt.withdraw()
    tk.messagebox.showinfo("App Edificio Guanabara",
                           "Ok, competência lançada com sucesso!")

# funcao que fecha a janela toplevel
def fechar_janela(janela):
    janela.destroy()

# função que grava o saldo do morador na tabela
def lanca_saldo(nomecondomino, saldo):
    saldo = float(saldo)
    novo_condominio_df = read_excel()
    lista_condominos = [novo_condominio_df.iloc[i, 0] for i in range(6)]

    for morador in lista_condominos:
        if nomecondomino in morador:
            # aqui vamos lançar o saldo do morador na tabela
            novo_condominio_df.loc[novo_condominio_df['Condômino']
                                   == morador, 'Saldo'] = neg(float(saldo))
    # e aqui exportamos para excel
    novo_condominio_df.to_excel('Calculos.xlsx')

    janela_finalizacao_saldo = tk.Tk()
    janela_finalizacao_saldo.withdraw()
    tk.messagebox.showinfo("App Edificio Guanabara",
                           "Ok, saldo lançado com sucesso!")


# lança uma despesa específica para um morador
def aplica_despesa_morador(morador, despesa, valor):
    valor = float(valor)

    novo_condominio_df = read_excel()
    novo_condominio_df.loc[novo_condominio_df['Condômino']
                           == morador, despesa] = valor
    novo_condominio_df.to_excel('Calculos.xlsx')

    janela_finalizacao_desp_espec = tk.Tk()
    janela_finalizacao_desp_espec.withdraw()
    tk.messagebox.showinfo("App Edificio Guanabara",
                           "Ok, despesa específica lançada com sucesso!")
    
# nesta funcao iremos gravar um histórico dos lançamentos de despesas e saldo de cada morador 
# no arquivo historico.txt antes de encerrar o programa, se o usuario desejar
def grava_historico():
    data_atual = datetime.now()
    # aqui faremos uma conversao do mes (numeral) para mes (por extenso)
    mes_ext = {1: 'jan', 2: 'fev', 3: 'mar', 4: 'abr', 5: 'mai', 6: 'jun',
               7: 'jul', 8: 'ago', 9: 'set', 10: 'out', 11: 'nov', 12: 'dez'}
    mes_atual = data_atual.strftime('%m/%Y')
    mes, ano = mes_atual.split("/")
    mes = '{}/{}'.format(mes_ext[int(mes)], ano)

    novo_condominio_df = read_excel()
    lista_condominos = [novo_condominio_df.iloc[i, 0] for i in range(6)]

    for morador in lista_condominos:
        
        saldo = novo_condominio_df.loc[novo_condominio_df['Condômino']
                                == morador, 'Saldo']
        saldo = float(saldo)       
        # caso haja algum saldo lançado, iremos gravar no histórico 
        if saldo != 0:        
            file = codecs.open("historico.txt", "a", encoding="cp1252")
            file.write(f'{morador} - Lançado saldo de {neg(saldo)} - mes: {mes}\n')
            file.close()

        lista_despesas = list(novo_condominio_df.columns)

        # aqui vamos repassar por todas as despesas da tabela (uma a uma)
        # e se houver alguma despesa (maior que zero) iremos escrever no bloco de notas
        n_despesa = 1
        while n_despesa <= 8:
            despesa = novo_condominio_df.loc[novo_condominio_df['Condômino']
                                                == morador, lista_despesas[n_despesa]].item()
            if despesa > 0:
                file = codecs.open("historico.txt", "a", encoding="cp1252")
                file.write(
                    f'{morador} - Despesa de {despesa} referente {lista_despesas[n_despesa]} - mês: {mes}\n')
                # aqui vamos subtrair as despesas do saldo
                saldo += float(despesa)
                file.close()
            n_despesa += 1
        # aqui iremos gravar no bloco de notas o saldo final do morador, caso ainda haja saldo
        if saldo < 0:
            file = codecs.open("historico.txt", "a", encoding="cp1252")
            file.write(f'{morador} - Saldo atualizado: {neg(saldo)} - mês: {mes}\n')
            file.close()
        file = codecs.open("historico.txt", "a", encoding="cp1252")
        file.write(f'\n')
        file.close()
    sys.exit()


def pegar_valor_copel():
    janela_escolha.update()
    # escondendo a tela do console do webdriver
    service = Service('chromedriver')
    service.creationflags = CREATE_NO_WINDOW

    driver = webdriver.Chrome(service=service)
    # os metodos abaixo irão acessar via web a área de cliente da Copel,
    # capturar o valor da unidade consumidora do condominio e enviar para a funcao aplica_despesa_geral
    driver.get("https://avanl.copel.com/avaweb/paginas/consultaDebitos.jsf")
    time.sleep(1)
    login = driver.find_element_by_xpath(
        '//*[@id="formulario:numDoc"]').send_keys('usuario')
    senha = driver.find_element_by_xpath('//*[@id="formulario:pass"]').click()
    time.sleep(2)
    senha = driver.find_element_by_xpath(
        '//*[@id="formulario:pass"]').send_keys('senha')
    time.sleep(1)
    driver.find_elements_by_tag_name('button')[0].click()
    driver.find_element_by_id('formLogin:tbUcs:1:j_idt50').click()
    time.sleep(3)
    driver.find_element_by_xpath(
        '//*[@id="sectiontemplate"]/div[2]/div/ul/li[3]/div/a/img').click()
    valor = driver.find_element_by_xpath(
        '//*[@id="j_idt48:0:dtDebitosUcLogada:0:j_idt72"]').text
    valor = valor.replace(',', '.')
    driver.close()
    aplica_despesa_geral('Energia Elétrica', valor)

# aplica uma despesa a ser rateado entre todos os moradores do condominio
def aplica_despesa_geral(despesa, valor):
    valor = float(valor)

    novo_condominio_df = read_excel()
    nova_despesa = novo_condominio_df.loc[novo_condominio_df['Condômino']
                                          == 'Totais', despesa] = valor
    ratio_despesa = math.ceil(nova_despesa / 6)
    lista_condominos = [novo_condominio_df.iloc[i, 0] for i in range(6)]
    for morador in lista_condominos:
        novo_condominio_df.loc[novo_condominio_df['Condômino']
                               == morador, despesa] = "{:.0f}".format(ratio_despesa)
    novo_condominio_df.to_excel('Calculos.xlsx')

    janela_finalizacao_desp_geral = tk.Tk()
    janela_finalizacao_desp_geral.withdraw()
    tk.messagebox.showinfo("App Edificio Guanabara",
                           "Ok, despesa geral lançada com sucesso!")

# essa funcao é responsavel por capturar o nr da tela que o usuario quer acessar
# essas telas (toplevel) por sua vez, irão trabalhar com as funcoes acima
def triagem():
    escolha = campo_comando.get()
    # opção que irá pegar o valor de consumo do condomonio da Copel
    if escolha == '1':
        pegar_valor_copel()
    # janela de lançamento de DESPESA GERAL
    elif escolha == '2':
        novo_condominio_df = read_excel()
        lista_despesas = list(novo_condominio_df.columns)
        lista_despesas = lista_despesas[1:9]

        janela_despesa_geral = tk.Toplevel()
        janela_despesa_geral.title('Despesa Geral')
        texto_label = Label(janela_despesa_geral, text="Escolha uma Despesa para Lançar",
                            borderwidth=2, relief='solid', fg='white', bg='black')
        texto_label.grid(row=0, column=0, padx=10, pady=10,
                         sticky='nsew', columnspan=3)

        campo_despesa_label = Label(janela_despesa_geral, text='Despesa')
        campo_despesa_label.grid(
            row=1, column=0, padx=10, pady=10, sticky='nsew')

        combobox_selecionar_despesa = ttk.Combobox(
            janela_despesa_geral, values=lista_despesas)
        combobox_selecionar_despesa.grid(
            row=1, column=1, padx=10, pady=15, sticky='nsew', ipadx=40, columnspan=2)

        campo_valor_label = Label(
            janela_despesa_geral, text='Valor da Despesa')
        campo_valor_label.grid(row=2, column=0, padx=10,
                               pady=10, sticky='nsew')

        campo_valor_despesa_geral = tk.Entry(janela_despesa_geral)
        campo_valor_despesa_geral.grid(
            row=2, column=1, padx=10, pady=10, sticky='nsew', columnspan=2)
        
        botao_lanca_despesa = tk.Button(
        janela_despesa_geral, text="Processar", bg='green', fg='white', relief='solid', command=lambda: aplica_despesa_geral(combobox_selecionar_despesa.get(), campo_valor_despesa_geral.get()))
        botao_lanca_despesa.grid(row=3, column=1, padx=3, pady=10, sticky='e', ipadx=30)
        
        botao_sair_despesa = tk.Button(
        janela_despesa_geral, text="Sair", command=lambda: fechar_janela(janela_despesa_geral))
        botao_sair_despesa.grid(row=3, column=2, padx=10, pady=10, sticky='nsew')  
    # janela de lançamento de DESPESA ESPECIFICA      
    elif escolha == '3':
        janela_despesa_especifica = tk.Toplevel()
        janela_despesa_especifica.title('Despesa Específica')
        texto_label = Label(janela_despesa_especifica, text="Escolha uma Despesa Específica para Lançar",
                            borderwidth=2, relief='solid', fg='white', bg='black')
        texto_label.grid(row=0, column=0, padx=10, pady=10,
                         sticky='nsew', columnspan=3)

        campo_morador = Label(janela_despesa_especifica, text='Morador')
        campo_morador.grid(
            row=1, column=0, padx=10, pady=10, sticky='nsew')

        novo_condominio_df = read_excel()
        lista_condominos = [novo_condominio_df.iloc[i, 0] for i in range(6)]
        
        lista_despesas = list(novo_condominio_df.columns)
        lista_despesas = lista_despesas[1:9]        

        combobox_selecionar_morador_desp_espec = ttk.Combobox(
            janela_despesa_especifica, values=lista_condominos)
        combobox_selecionar_morador_desp_espec.grid(
            row=1, column=1, padx=10, pady=10, sticky='nsew', columnspan=2)

        campo_despesa_label = Label(janela_despesa_especifica, text='Despesa')
        campo_despesa_label.grid(
            row=2, column=0, padx=10, pady=10, sticky='nsew')

        combobox_selecionar_despesa_espec = ttk.Combobox(
            janela_despesa_especifica, values=lista_despesas)
        combobox_selecionar_despesa_espec.grid(
            row=2, column=1, padx=10, pady=10, sticky='nsew', columnspan=2)

        campo_valor_label = Label(
            janela_despesa_especifica, text='Valor da Despesa')
        campo_valor_label.grid(row=3, column=0, padx=10,
                               pady=10, sticky='nsew')

        campo_valor_desp_espec = tk.Entry(janela_despesa_especifica)
        campo_valor_desp_espec.grid(
            row=3, column=1, padx=10, pady=10, sticky='nsew', ipadx=50, columnspan=2)

        botao_lanca_despesa_espec = tk.Button(
        janela_despesa_especifica, text="Processar", bg='green', fg='white', relief='solid', command=lambda: aplica_despesa_morador(combobox_selecionar_morador_desp_espec.get(), combobox_selecionar_despesa_espec.get(), campo_valor_desp_espec.get()))
        botao_lanca_despesa_espec.grid(row=4, column=1, padx=3, pady=10, sticky='e', ipadx=30)
        
        botao_sair_despesa = tk.Button(
        janela_despesa_especifica, text="Sair", command=lambda: fechar_janela(janela_despesa_especifica))
        botao_sair_despesa.grid(row=4, column=2, padx=10, pady=10, sticky='nsew')      
    # janela de lançamento do SALDO DO MORADOR    
    elif escolha == '4':
        janela_saldo = tk.Toplevel()
        janela_saldo.title('Saldo para Morador')
        texto_label = Label(janela_saldo, text="Escolha uma Morador para Lançar Saldo",
                            borderwidth=2, relief='solid', fg='white', bg='black')
        texto_label.grid(row=0, column=0, padx=10, pady=10,
                         sticky='nsew', columnspan=3)

        campo_morador_saldo = Label(janela_saldo, text='Morador')
        campo_morador_saldo.grid(
            row=1, column=0, padx=10, pady=10, sticky='nsew')

        novo_condominio_df = read_excel()
        lista_despesas = list(novo_condominio_df.columns)
        lista_condominos = [novo_condominio_df.iloc[i, 0] for i in range(6)]

        combobox_selecionar_morador_saldo = ttk.Combobox(
            janela_saldo, values=lista_condominos)
        combobox_selecionar_morador_saldo.grid(
            row=1, column=1, padx=10, pady=10, sticky='nsew', ipadx=40, columnspan=2)

        campo_valor_label = Label(janela_saldo, text='Valor do Saldo')
        campo_valor_label.grid(row=3, column=0, padx=10,
                               pady=10, sticky='nsew')

        campo_valor_saldo = tk.Entry(janela_saldo)
        campo_valor_saldo.grid(row=3, column=1, padx=10,
                               pady=10, sticky='nsew', columnspan=2)

        botao_lanca_saldo = tk.Button(
        janela_saldo, text="Processar", bg='green', fg='white', relief='solid', command=lambda: lanca_saldo(combobox_selecionar_morador_saldo.get(), campo_valor_saldo.get()))
        botao_lanca_saldo.grid(row=4, column=1, padx=3, pady=10, sticky='e', ipadx=30)
        
        botao_sair_saldo = tk.Button(
        janela_saldo, text="Sair", command=lambda: fechar_janela(janela_saldo))
        botao_sair_saldo.grid(row=4, column=2, padx=10, pady=10, sticky='nsew')        
    elif escolha == '5':
        define_competencia()
    elif escolha == '6':
        limpar_linhas()
    elif escolha == '7':
        ao_fechar()


# Janela Principal com o Menu de Opções
janela_escolha = tk.Tk()
janela_escolha.title("Edificio Guanabara")

despesas = ['1. Pegar Valor Copel', 
            '2. Lançar Despesa Geral',
            '3. Lançar Despesa Específica',
            '4. Lançar Saldo para Morador', 
            '5. Lançar Competência (automático)', 
            '6. Limpar Planilha ',
            '7. Sair']

texto_label = Label(janela_escolha, text="Escolha uma das opções abaixo",
                    relief='solid', fg='white', bg='black')
texto_label.grid(row=0, column=0, padx=10, pady=5,
                 sticky='nsew', columnspan=2)
linha = 1
for i in despesas:
    linha_escolha = Label(janela_escolha, text=i)
    linha_escolha.grid(row=linha, column=0, padx=0,
                       pady=0, ipadx=30, sticky='w')
    linha += 1
campo_comando = tk.Entry(janela_escolha)
campo_comando.grid(row=linha, column=0, padx=10, pady=3, sticky='nsew', columnspan=2)
linha += 1

botao_continuar = tk.Button(
    janela_escolha, text="Processar", command=triagem, bg='green', fg='white', relief='solid')
botao_continuar.grid(row=9, column=0, padx=3, pady=10, sticky='e', ipadx=25)

botao_sair = tk.Button(
    janela_escolha, text="Sair", command=ao_fechar)
botao_sair.grid(row=linha, column=1, padx=10, pady=10, sticky='nsew', ipadx=20)
janela_escolha.protocol("WM_DELETE_WINDOW", ao_fechar)

janela_escolha.mainloop()
