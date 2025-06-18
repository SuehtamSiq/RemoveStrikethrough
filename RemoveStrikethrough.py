from docx import Document
import re

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

def alterarSimboloOrdinal(doc):
    # Define o padrão para buscar números seguidos de "o", "a", "s", etc.
    padrao_ordinais = r'(\d+)(o|a|s)(?=\s|$)'

    for paragrafo in doc.paragraphs:
        for run in paragrafo.runs:
            if run.font.strike:
                continue  # ignora texto tachado
            # Substitui o padrão de número ordinal com o símbolo '°'
            texto_corrigido = re.sub(padrao_ordinais, r'\1°', run.text)
            run.text = texto_corrigido

removerTextoTachado('C:\\Matheus Siqueira\\Python\\RemoveStrikethrough\\LEI DE REGISTROS PÚBLICOS (TESTE).docx', 
                    'C:\\Matheus Siqueira\\Python\\RemoveStrikethrough\\saida_sem_tachado.docx')
