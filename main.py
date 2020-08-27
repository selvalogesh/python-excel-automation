import os, sys
from automation import ExcelToPdf  as E2P

input_directory = "./input"
output_directory = "./output"

if __name__ == "__main__":
    if(os.path.isdir(input_directory) and os.path.isdir(output_directory)):
        excelFiles = [ os.path.join(input_directory, file_name).replace("\\","/") for file_name in os.listdir(input_directory) if file_name.endswith(".xlsx")]
    else:
        print("Input or Output directory (./input or ./output) cannot be found.")
        sys.exit(1)

    for excelFile in excelFiles:
        pdfOutputDir = os.path.join(output_directory, os.path.basename(excelFile).replace(".xlsx","")).replace("\\","/")
        os.makedirs(pdfOutputDir, exist_ok=True)
        with open(os.path.join(pdfOutputDir, "error.txt"), 'w') as fp: 
            pass
        E2P(excelFile, pdfOutputDir)