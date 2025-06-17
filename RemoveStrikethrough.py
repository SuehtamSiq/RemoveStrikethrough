from docx import Document

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
            
    doc.save(caminho_saida)
    print(f'Documento salvo sem texto tachado em: {caminho_saida}')

removerTextoTachado('entrada.docx', 'saida_sem_tachado.docx')