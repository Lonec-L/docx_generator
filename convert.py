import argparse
import pathlib
import sys
import os
import comtypes.client
from tqdm import tqdm

wdFormatPDF = 17

directory = "output"

parser = argparse.ArgumentParser(
    prog="Utility script that complements docx_generator. Takes all docx files in output folder and converts them to pdf (output is stored in separate ./pdf folder). Additionally it can combine all converted pdf files into one and remove generated pdf files from pdf folder.",
    description="A simple python script that allows batch generating docx files from template and keys and values stored in csv file.",
)

parser.add_argument("-c", "--combine", help="Combines all converted files into one pdf (order is not guaranteed, tho it seems like they are included in alphabetical order)", action='store_true')

parser.add_argument("-d", "--delete", help="Deletes genereated pdf files after conversion. Has no effect if used without --combine",action='store_true')
args = parser.parse_args()



path =  pathlib.Path("./pdf")
path.mkdir(exist_ok=True)
script_dir = os.path.dirname(__file__)

directory_path = os.path.join(script_dir, directory)

pdf_path = os.path.join(script_dir, "pdf")

print("Opening Word client...")
word = comtypes.client.CreateObject('Word.Application')

print("Converting docx to pdf:")
for filename in tqdm(os.listdir(directory)):
    f = os.path.join(directory_path, filename)
    # checking if it is a file
    if os.path.isfile(f):
        doc = word.Documents.Open(f)
        doc.SaveAs(pdf_path+"/"+filename+".pdf", FileFormat=wdFormatPDF)
        doc.Close()
word.Quit()

if(args.combine):
    print("Combining pdf files into one...")
    if(args.delete):
        print("Deleting pdf files after combine...")
    from pypdf import PdfWriter

    merger = PdfWriter()

    for filename in tqdm(os.listdir("pdf")):
        f = os.path.join("pdf", filename)
        if(os.path.isfile(f)):
            merger.append(f)
            if(args.delete):
                os.remove(f)
    if(args.delete):
        os.rmdir("pdf")


    merger.write("merged-pdf.pdf")
    merger.close()
print("Done!")