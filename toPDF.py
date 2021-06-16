import os
import win32com.client
import glob
import pathlib
import PyPDF2
import shutil
from datetime import datetime

input("実行するにはエンターを押してください！")
dt_now = datetime.now()
print(dt_now)
print("実行中")


# wordをpdfに変換する関数


def WDtoPDF(in_wd, out_wd, formatType=17):
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(in_wd)
    doc.SaveAs(out_wd, formatType)
    doc.Close()
    word.Quit()

# pdfをまとめて１つにする関数


def pdf_merger(out_pdf, pdfs):
    merger = PyPDF2.PdfFileMerger()

    for pdf in pdfs:
        print(pdf)
        merger.append(pdf)

    merger.write(out_pdf)
    merger.close()


# 結合するexcel,word,powerpoint,pdfがある場所を自動取得
file_path = os.getcwd()

# 変換したPDFファイル・結合したPDFファイルを入れるフォルダー
sub_name = 'PDF'

file0 = pathlib.Path(file_path)
file1 = file0.joinpath(sub_name)

# フォルダーを作成
os.makedirs(file1, exist_ok=True)

# もとのファイルがある場所に移動
os.chdir(file_path)
files = glob.glob('*')


# excel,word,powerpointをpdfに変換し、名前の拡張子も.pdfに変更
# 変換したPDFのパスをリストに渡す
pdfs = []

for f in files:
    file_p = file0.joinpath(f)

    if file_p.suffix == '.docx':
        file_pdf = file1.joinpath(f.replace(file_p.suffix, '.pdf'))
        WDtoPDF(str(file_p), str(file_pdf))
        pdfs.append(str(file_pdf))

    elif file_p.suffix == '.pdf':
        shutil.copy('./' + f, './' + sub_name)  # もともとPDFであるものはそのままコピー
        file_pdf = file1.joinpath(f)
        pdfs.append(str(file_pdf))
    else:
        pass

# 結合したPDFに名前を付け、フォルダーに置く
out_file = str(pathlib.Path(file1).joinpath('out.pdf'))
# PDFの結合
pdf_merger(out_file, pdfs)

input("実行完了！エンターで終了します")
