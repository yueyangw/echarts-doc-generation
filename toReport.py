import base64, json
import os

from flask import Flask, request, send_from_directory
from docx import Document
from docx.shared import Cm, Pt
from io import BytesIO
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)


@app.route("/")
def mainPage():
    html = open("pie-roseType-simple.html", "r")
    return html.read()


@app.route("/savedoc", methods=['POST', 'GET'])
def saveAsDocument():
    if request.method == 'POST':
        datas = request.get_data()
        datas = json.loads(datas)
        img = base64ToImage(datas['imgdata'])
        template = Document('./template.docx')
        doc = docBuilder(template, img, datas['echartsData'])
        doc.save(r'report.docx')
        dic = os.getcwd()
        res = send_from_directory(dic, 'report.docx', as_attachment=True)
        return res


def base64ToImage(imgdata):
    strs = imgdata[22:]  # 返回的base64字符串前面有几个多余字符，这里删掉
    imgdata = base64.b64decode(strs)  # 把base64的图片解码成二进制返回
    return imgdata


def docBuilder(doc, img, datas):
    basicInfo = datas['series'][0]
    mapp = {
        'pie': '饼图'
    }
    datas = {
        'topic': basicInfo['name'],
        'picture': img,
        'category': mapp[basicInfo['type']],
        'data': basicInfo['data']
    }
    replaceText(doc, "<topic>", datas['topic'])
    replaceText(doc, "<picture>", img)
    replaceText(doc, "<category>", datas['category'])
    return doc


def replaceText(doc, oldText, newText):
    for para in doc.paragraphs:
        if oldText in para.text:
            if oldText == '<picture>':
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para.text = para.text.replace(oldText, 'pppiiictturree')
                for run in para.runs:
                    if 'pppiiictturree' in run.text:
                        run.text = run.text.replace('pppiiictturree', '')
                        run.add_picture(BytesIO(newText), width=Cm(14))
            else:
                para.text = para.text.replace(oldText, newText)
                for run in para.runs:
                    if oldText == '<topic>':
                        run.font.name = '宋体'
                        run.font.size = Pt(27)
                    else:
                        run.font.name = '楷体'
                        run.font.size = Pt(16)