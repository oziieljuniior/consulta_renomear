import tkinter as tk
from tkinter import filedialog
import os

import numpy as np
import pandas as pd

print("Bem-Vindo")
print("Selecione a planilha salva da consulta")

planilha = filedialog.askopenfilename()
print(planilha)
data1 = pd.DataFrame(pd.read_excel(planilha))
#print(data1)
print("*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*")
print("Selecione o caminho onde os arquivos estão salvos")
caminho = filedialog.askdirectory()
print(caminho)
lista_caminho = os.listdir(caminho)
lista_caminho.sort()
t = len(lista_caminho)
print("*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*")
print("Selecione o caminho onde os arquivos devem ficar salvos")
salvar = filedialog.askdirectory()
print(salvar)
print("*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*")
print("Relalizando conversões")

data2 = np.array(data1['codigo'])
data3 = np.array(data1['Ano$Edicao'])
a = []
b = []
e = []
for i in range(0,t):
    c = lista_caminho[i].split('_')
    c[1] = c[1].replace('.pdf','')
    a.append(c[0])
    b.append(c[1])

    print(type(b[i]),b[i])
print("*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*")
print("Relalizando conversões")

for i in range(0,t):
    if b[i] == str(data2[i,]):
        c = salvar + "/" + a[1] + "_" + data3[i,] + ".pdf"
        e.append(c)
    else:
        continue
#print(e)

print("*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*")
print("Relalizando conversões")
j = 0
for i in range(0,t):
    ###------Renomear
    #
    #/home/oziel/Documentos/teste1/Mercurio peruano, Lima_828531196.pdf
    antigo_nome = caminho + "/" + lista_caminho[i]
#    i = str(i)
    novo_nome = e[i]
#    i = int(i)
    #antigo_nome = lista_caminho[i]
    #novo_nome = a[i] +"_" + titulo[0] + '.pdf'
    
    print(antigo_nome)
    print(novo_nome)
    try:
        os.rename(antigo_nome ,novo_nome)
    except FileNotFoundError:
        print("Ok, algum erro aconteceu. Mas vou continuar a conversão. No final vou mostrar quantos arquivos foram convertidos")
        j += 1
        continue
#    i += 1
    
print("*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*")
print("Conversão realizado com sucesso!")
print("Pode ser que ainda tenha arquivos que não foram convertidos.")
print("j = ", j)
print("Se j = 0. Então ignore")
print("Caso contrario, contate o autor do código")    