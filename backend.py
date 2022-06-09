from __future__ import print_function
from sys import platform
import time
from flask import Flask, render_template, request, url_for, flash, redirect
from mailmerge import MailMerge
import uno
import pdfkit
from comtypes.client import CreateObject
wdFormatPDF = 17

config = pdfkit.configuration(wkhtmltopdf='/usr/bin/wkhtmltopdf')
options = {'--load-error-handling': 'ignore'}

template_docx = "docs/andrahandansokan2.docx"

template_odt = "docs/andrahandansokan.odt"
document = MailMerge(template_docx)
print(document.get_merge_fields())

app = Flask(__name__)
app.config['SECRET_KEY'] = '3f04e3c6ff42ee4b1f44f4f782ec9f7c'


@app.route("/")
@app.route("/home")
def home():
    return render_template('home.html')


@app.route("/andrahandansokan", methods=['POST', 'GET'])
def ansokan():
    return render_template('andrahandansokan.html')


@app.route("/success_ansok", methods=['POST', 'GET'])
def success_submit():
    if request.method == "POST":
        # print(request.form)
        docx_out_name = "andrahandansokan_lgh_{0}_{1}.docx".format(
            request.form['lgh_nmr'], request.form['datum'])
        odt_out_name = "andrahandansokan_lgh_{0}_{1}.odt".format(
            request.form['lgh_nmr'], request.form['datum'])
        pdf_out_name = "andrahandansokan_lgh_{0}_{1}.pdf".format(
            request.form['lgh_nmr'], request.form['datum'])

        document.merge(landper_fname=request.form['landper_fname'],
                       landper_lname=request.form['landper_lname'],
                       lgh_nmr=request.form['lgh_nmr'],
                       trp=request.form['trp'],
                       landper_mblnr=request.form['landper_mblnr'],
                       landper_email=request.form['landper_email'],
                       startdate=request.form['startdate'],
                       enddate=request.form['enddate'],
                       landper_co_fname=request.form['landper_co_fname'],
                       landper_co_lname=request.form['landper_co_lname'],
                       landper_ny_adress=request.form['landper_ny_adress'],
                       landper_ny_postn=request.form['landper_ny_postn'],
                       landper_ny_ort=request.form['landper_ny_ort'],
                       skal=request.form['skal'],
                       tnt_fname=request.form['tnt_fname'],
                       tnt_lname=request.form['tnt_lname'],
                       tnt_mblnr=request.form['tnt_mblnr'],
                       tnt_email=request.form['tnt_email'],
                       tnt_arbtgvr=request.form['tnt_arbtgvr'],
                       tnt_arbgvr_mblnr=request.form['tnt_arbgvr_mblnr'],
                       nuv_hyrsgst=request.form['nuv_hyrsgst'],
                       undrs_ort=request.form['undrs_ort'],
                       datum=request.form['datum'])
        document.write(docx_out_name)
        time.sleep(5)
        if platform == 'win32':
            word = CreateObject('Word.Application')
            doc = word.Documents.Open(docx_out_name)
            doc.SaveAs(pdf_out_name, FileFormat=wdFormatPDF)
            doc.Close()
            word.Quit()

        if platform == 'linux' or platform == 'linux2':
            pdfkit.from_file(docx_out_name, pdf_out_name,
                             configuration=config, options=options)

        # print(landper_fname)
        # print(landper_lname)
        # print(lgh_nmr)
        # print(trp)
        # print(landper_mblnr)
        # print(landper_email)
        # print(startdate)
        # print(enddate)
        # print(landper_co_fname)
        # print(landper_co_lname)
        # print(landper_ny_adress)
        # print(landper_ny_postn)
        # print(landper_ny_ort)
        # print(skal)
        # print(tnt_fname)
        # print(tnt_lname)
        # print(tnt_mblnr)
        # print(tnt_email)
        # print(tnt_arbtgvr)
        # print(tnt_arbgvr_mblnr)
        # print(nuv_hyrsgst)
        # print(underskrift)
        # print(undrs_ort)
        # print(datum)

        # flash(f'Din ansökan skickades', 'success')
        return redirect(url_for('ansokan'))


if __name__ == '__main__':
    app.run(debug=True)
