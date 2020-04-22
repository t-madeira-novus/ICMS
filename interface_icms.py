from appJar import gui
import pandas as pd
import calcular_icms as ci
import banco_dados_icms as bd

def _pegar_relatorio():
    global nome_arquivo
    nome_arquivo = app.openBox(title=None, dirName=None, fileTypes=None, asFile=False, parent=None, multiple=True, mode='r')

def _thread_calcular_icms():
    global nome_arquivo

    path_to_save = app.saveBox(title="", fileName="",
                       dirName=None, fileExt=".csv", fileTypes=None, asFile=None, parent=None)

    app.thread(ci._calcular(nome_arquivo, path_to_save))

def _thread_popular_banco():
    app.thread(bd._popular_banco(nome_arquivo))

# Variaveis globais
nome_arquivo = ""

# Criando a interface Gráfica
app = gui("iCmS")
app.setFont(10)

#######################################################################################################################
coluna = 0
linha = 0

app.addButton("Pegar relatório gerado pela Domínio", _pegar_relatorio, row = coluna, column = linha)
linha += 1
app.addButton("Calcular ICMS", _thread_calcular_icms, row = coluna, column = linha)
linha += 1
app.addButton("Popular banco de dados", _thread_popular_banco, row = coluna, column = linha)
linha += 1

########################################################################################################################


# start the GUI
app.go()