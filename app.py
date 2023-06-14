from funcoes import *

# Criando uma instância da classe planilha_automatica
planilha = planilha_automatica()

# Chamando os métodos na instância planilha
apresentacao()
planilha.criando_paginas()
while True:
    planilha.escolher_sheet()
    planilha.inserir_colunas()
    planilha.inserir_conteudo()
    continuar = input('deseja adicionar conteudo a outra pagina? (sim/nao)')
    if continuar != 'sim':
        break
planilha.salvar_wb()
