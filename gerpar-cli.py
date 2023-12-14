# -*- coding: utf-8 -*-
from docx import Document
from openpyxl import load_workbook

from gerpar import parecer_edital, parecer_contrato

def make_parecer(data):
    assunto = data[0]
    processo = data[1]
    num_modalidade = data[2]
    dataEdital = data[3]
    dataContrato = data[4]

    # São gerados os pareceres
    parecer_edital(assunto, processo, num_modalidade, dataEdital)
    parecer_contrato(assunto, processo, num_modalidade, dataContrato)


def main():
    wb = load_workbook(filename='relacao_contratos.xlsx')
    ws = wb.active
    for row in ws.iter_rows(values_only=True):
        make_parecer(row)

if __name__ == "__main__":
   main()