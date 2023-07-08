import os
import pandas as pd
from pandas_datareader.data import get_quote_yahoo
import datetime
import math
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.pagesizes import A4, portrait #用紙の向き
from reportlab.platypus import Table, TableStyle
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus.frames import Frame

import locale
locale.setlocale(locale.LC_CTYPE, "Japanese_Japan.932")

t_delta = datetime.timedelta(hours=9)
JST = datetime.timezone(t_delta, "JST")
now = datetime.datetime.now(JST)
d = now.strftime("%Y%m%d")
ymd = now.strftime("%Y年%m月%d日")
tax = 0.1 # 税率
inputPath = "./input"
## 日本円米ドル 為替
# result = get_quote_yahoo('USDJPY=X')
# aryResult = result["price"].values
# moneyOrder = aryResult[0]
moneyOrder = input("現在のドル円を入力してください\n>")

def make(pdfData, linodeCSVPath):
    filename=linodeCSVPath.replace("./input/", "").replace(".csv", "") #ファイル名
    print("Making PDF " + filename + " ...")
    pdfCanvas = setInfo(filename) #キャンバス名
    printString(pdfCanvas, pdfData)
    pdfCanvas.save() #保存

def setInfo(filename):
    pdfCanvas = canvas.Canvas("./{0}_{1}.pdf".format(d,filename)) #保存先
    pdfCanvas.setAuthor("kuwa") #作者
    pdfCanvas.setTitle("linode-price")  #表題
    pdfCanvas.setSubject("")  #件名
    return pdfCanvas

def printString(pdfCanvas, pdfData):
    
    maxlow = 45 #行数
    pages = (len(pdfData) + 1) / maxlow #ページ数
    num = math.ceil(pages)
        
    for i in range(0, len(pdfData), (maxlow-1)):
        tmpPdfData = pdfData[i: i+(maxlow-1)]
        pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5")) #フォント
        width, height = A4 #用紙サイズ
        font_size = 24 #フォントサイズ

        # タイトル
        pdfCanvas.setFont("HeiseiKakuGo-W5", font_size)
        pdfCanvas.drawString(230, 770, "Linode請求内訳") #書き出し(横位置, 縦位置, 文字)

        # 作成日
        font_size = 7 #フォントサイズ
        pdfCanvas.setFont("HeiseiKakuGo-W5", font_size)
        pdfCanvas.drawString(465, 770,  f"作成日: {ymd}")


        # ノードと価格
        detailData = [
                ["Node","価格(USドル)","税金(USドル)"],
            ]

        # priceData = [row[0:2].extend(row[5:5]).append(row[5:5]*0.1) for row in tmpPdfData]
        totalPrice = 0
        priceData = []
        for row in tmpPdfData: # メインの価格が書いてあるところを追加していく
            priceData.append([row[0], row[5], float(row[5]) * tax])
            totalPrice += ((float(row[5]) * tax) + float(row[5]))

        detailData.extend(priceData)

        if len(detailData) <= maxlow:
            index_no = maxlow - len(detailData)

            for idx in range(index_no):
                detailData.append([" ", " ", " "])

        table = Table(detailData, colWidths=(105*mm,35*mm,35*mm), rowHeights=5*mm)
        table.setStyle(TableStyle([
                ("FONT", (0, 0), (-1, -1), "HeiseiKakuGo-W5", font_size),#フォント
                ("BOX", (0, 0), (-1, -(index_no + 1)), 1, colors.black),#罫線
                ("INNERGRID", (0, 0), (-1, -(index_no + 1)), 1, colors.black),#罫線
                ("ALIGN", (1, 0), (6, -1), "RIGHT"),#右揃え
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),#フォント位置
            ]))
        table.wrapOn(pdfCanvas, 20*mm, 40*mm)#table位置
        table.drawOn(pdfCanvas, 20*mm, 40*mm)

        header_data = [
                ["合計価格(税込み USドル) ", totalPrice],
                ["合計価格(税込み 日本円) ", round(totalPrice * float(moneyOrder))]
            ]

        table = Table(header_data, colWidths=(35*mm,35*mm), rowHeights=5*mm)#tableの大きさ
        table.setStyle(TableStyle([#tableの装飾
                ("FONT", (0, 0), (-1, -1), "HeiseiKakuGo-W5", font_size),#フォント
                ("BOX", (0, 0), (-1, -1), 1, colors.black),#罫線
                ("ALIGN", (1, 0), (6, -1), "RIGHT"),#右揃え
                ("INNERGRID", (0, 0), (-1, -1), 1, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),#フォント位置
            ]))
        table.wrapOn(pdfCanvas, 125*mm, (250-(len(priceData)*5))*mm)#table位置(0を一番下からとして計算 各行は5mm)
        table.drawOn(pdfCanvas, 125*mm, (250-(len(priceData)*5))*mm)

        ## 1枚目終了(改ページ)
        pdfCanvas.showPage()


if __name__ == "__main__":
    files = os.listdir(inputPath)
    linodeCSVsPath = [inputPath+"/"+f for f in files if os.path.isfile(os.path.join(inputPath, f))]

    for linodeCSVPath in linodeCSVsPath:
        print("Loading " + linodeCSVPath + " ...")
        
        #linodeCSV読み込み
        dfCsv = pd.read_csv(filepath_or_buffer=linodeCSVPath, encoding="utf-8")
        #PDF作成
        make(dfCsv.values.tolist(), linodeCSVPath)
        print("Complete!")
    input()