import os
from flask import Flask, render_template, redirect, request, send_file, url_for

from gerpar import parecer_edital, parecer_contrato

app = Flask(__name__)


@app.route('/')
def index():
    return render_template("index.html")


@app.route('/gerpar', methods=['POST'])
def gerpar():
    # REcebe os dados dos pareceres a serem gerados
    requerente = request.form['requerente']
    assunto = request.form['assunto']
    processo = request.form['processo']
    num_modalidade = request.form['num_modalidade']
    dataEdital = request.form['dataEdital']
    dataContrato = request.form['dataContrato']

    # São gerados os pareceres
    parecer_edital(requerente, assunto, processo, num_modalidade, dataEdital)
    parecer_contrato(requerente, assunto, processo, num_modalidade, dataContrato)

    # Nomes dos pareceres para serem exibidos na pagina download
    nome_parecer_edital = ("Parecer_Edital_%s.docx" % (processo))
    nome_parecer_contrato = ("Parecer_Contrato_%s.docx" % (processo))

    return render_template("downloads.html", parecer_edital=nome_parecer_edital, parecer_contrato=nome_parecer_contrato)


@app.route('/download/<arquivo>', methods=["POST", "GET"])
def download(arquivo):
    # Para criar o link de donload
    base_dir = os.path.dirname(__file__)
    para_donwload = ("%s/pareceres/%s" % (base_dir, arquivo))

    # Excecao para o download do arquivo do parecer
    try:
        return send_file(para_donwload, as_attachment=True)
    except Exception as e:
        return redirect("/")

