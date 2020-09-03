import os
import time
import requests
import pandas as pd
from PIL import Image

def ExcelToPdf(excelFile, pdfOutputDir):
    s = requests.Session()
    dfs = pd.read_excel(excelFile, sheet_name="SheetJS")
    for row in range(len(dfs)-1):
        imagelist = []
        firstImg = None
        name = dfs["Student Name"][row]
        roll = dfs["Roll No./Register No"][row]

        for column in dfs:
            if (column.find("-") != -1):
                try:
                    urlRow = dfs[column][row]
                    urlRow = urlRow.split()[-1::-1]
                    urls = [x.strip() for x in urlRow]
                    for url in urls:
                        response = s.get(url, stream=True)
                        #print(response.status_code)
                        if(response.status_code == 200):
                            img = Image.open(response.raw)
                            if(img.mode == 'RGBA'):
                                img = img.convert('RGB')
                            if(firstImg == None):
                                firstImg = img
                            else:
                                imagelist.append(img)
                        else:
                            NoneImg = Image.open("None.jpg")
                            imagelist.append(NoneImg)
                except:
                    errorLog = os.path.join(pdfOutputDir,"error.txt").replace("\\","/")
                    with open(errorLog,'a') as err:
                        err.write("{0}_{1}_{2} in {3} -> {4}\n".format(row, roll, name, column, dfs[column][row]))
                    print("Error in {0}_{1}_{2} in {3} -> {4}".format(row, roll, name, column, dfs[column][row]))
                    NoneImg = Image.open("None.jpg")
                    imagelist.append(NoneImg)
        if(firstImg == None):
            firstImg = Image.open("None.jpg")
        pdfFilePath = os.path.join(pdfOutputDir,"{0}_{1}_{2}.pdf".format(row, roll, name)).replace("\\","/")
        firstImg.save(pdfFilePath,save_all=True, append_images=imagelist)
        print("Done {0}_{1}_{2}".format(row, roll, name))