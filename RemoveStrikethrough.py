from docx import Document
import re
import os

def removerTextoTachado(caminho_entrada, caminho_saida):
    doc = Document(caminho_entrada)

    for paragrafo in doc.paragraphs:
        novas_runs = []
        for run in paragrafo.runs:
            if run.font.strike:
                continue  # ignora texto tachado
            novas_runs.append(run.text)

        paragrafo.clear()  # limpa o parágrafo atual

        for texto in novas_runs:
            paragrafo.add_run(texto)  # adiciona texto não tachado

    # Chama a função para alterar os números ordinais
    alterarSimboloOrdinal(doc)

    doc.save(caminho_saida)
    print(f'Documento salvo sem texto tachado em: {caminho_saida}')
    
    # Corrigindo para tratar o caminho com espaços corretamente
    os.system(f'start "" "{caminho_saida}"')


def alterarSimboloOrdinal(doc):
    # Atualiza o padrão para capturar números ordinais simples ou seguidos de sufixos alfabéticos (como -B, -C)
    padrao_ordinais = r'(\d+)(o|a|s)(-[A-Za-z])?(?=\s|$)'  # Aceita ordinais como 3o, 4o-B, 5a-C, etc.

    for paragrafo in doc.paragraphs:
        for run in paragrafo.runs:
            if run.font.strike:
                continue  # ignora texto tachado
            # Substitui o padrão de número ordinal com o símbolo '°'
            texto_corrigido = re.sub(padrao_ordinais, r'\1°\3', run.text)  # Coloca o símbolo '°' e preserva o sufixo, se houver
            run.text = texto_corrigido

# Chamada de função com o caminho corrigido
removerTextoTachado(r'C:\Matheus Siqueira\Python\RemoveStrikethrough\LEI DE REGISTROS PÚBLICOS (TESTE).docx', 
                    r'C:\Matheus Siqueira\Python\RemoveStrikethrough\saida_sem_tachado.docx')

