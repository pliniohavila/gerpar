# -*- coding: utf-8 -*-
from docx import Document

# Para criar o parecer do edital


def parecer_edital(assunto, processo, num_modalidade, data):

    # Abre o documento modelo de base
    with open("modelos/mp_edital.docx", "rb") as docFile:
        doc = Document(docFile)

    # Laco for que pecorre o documento e realiza a adicao
    # das informacoes do novo parecer
    for paragraph in doc.paragraphs:
    #     if "[REQUERENTE]" in paragraph.text:
    #         orig_text = paragraph.text
    #         new_text = str.replace(orig_text, "[REQUERENTE]", requerente)
    #         paragraph.text = new_text

        if "[ASSUNTO]" in paragraph.text:
            for run in paragraph.runs:
                if "[ASSUNTO]" in run.text:
                    run.text = run.text.replace("[ASSUNTO]", assunto)
                    break
            # orig_text = paragraph.text
            # new_text = str.replace(orig_text, "[ASSUNTO]", assunto)
            # paragraph.text = new_text

        elif "[MODALIDADE_N]" in paragraph.text:
            orig_text = paragraph.text
            new_text = str.replace(orig_text, "[MODALIDADE_N]", num_modalidade)
            paragraph.text = new_text

        elif "[PROCESSO_N]" in paragraph.text:
            orig_text = paragraph.text
            new_text = str.replace(orig_text, "[PROCESSO_N]", processo)
            paragraph.text = new_text
      
        elif "[DATA]" in paragraph.text:
            orig_text = paragraph.text
            new_text = str.replace(orig_text, "[DATA]", data)
            paragraph.text = new_text

    # Salva o novo parecer criado
    nome_documento = str.replace(processo, '/', '-')
    doc.save("pareceres/Parecer_Edital_%s.docx" % (nome_documento))

# Para criar o parecer do contrato


def parecer_contrato(assunto, processo, num_modalidade, data):

    # Abre o documento modelo de base
    with open("modelos/mp_contrato.docx", "rb") as docFile:
        doc = Document(docFile)

    # Laco for que pecorre o documento e realiza a adicao
    # das informacoes do novo parecer
    for paragraph in doc.paragraphs:
    #     if "[REQUERENTE]" in paragraph.text:
    #         orig_text = paragraph.text
    #         new_text = str.replace(orig_text, "[REQUERENTE]", requerente)
    #         paragraph.text = new_text
        if "[ASSUNTO]" in paragraph.text:
            for run in paragraph.runs:
                if "[ASSUNTO]" in run.text:
                    run.text = run.text.replace("[ASSUNTO]", assunto)
                    break
        elif "[MODALIDADE_N]" in paragraph.text:
            orig_text = paragraph.text
            new_text = str.replace(orig_text, "[MODALIDADE_N]", num_modalidade)
            paragraph.text = new_text

        elif "[PROCESSO_N]" in paragraph.text:
            orig_text = paragraph.text
            new_text = str.replace(orig_text, "[PROCESSO_N]", processo)
            paragraph.text = new_text

        elif "[DATA]" in paragraph.text:
            orig_text = paragraph.text
            new_text = str.replace(orig_text, "[DATA]", data)
            paragraph.text = new_text

    # Salva o novo parecer criado
    nome_documento = str.replace(processo, '/', '-')
    doc.save("pareceres/Parecer_Contrato_%s.docx" % (nome_documento))
