import base64, os, json
from flask import Flask, url_for, request, Response
from docx import Document

app = Flask(__name__)

def base64ToImage(imgdata):
    strs = imgdata[22:] #返回的base64字符串前面有几个多余字符，这里删掉
    imgdata = base64.b64decode(strs) #把base64的图片解码成二进制返回
    return imgdata

@app.route("/")
def mainPage():
    html = open("pie-roseType-simple.html", "r")
    return html.read()

@app.route("/saveimg", methods=['POST', 'GET'])
def saveAsImage():
    if request.method == 'POST':
        datas = request.get_data()
        datas = json.loads(datas)
        img = base64ToImage(datas['imgdata'])
        res = Response(img, mimetype="image/png")
        return res