import copy
from docx import Document
from python_docx_replace import docx_replace

import argparse
import pathlib

from tqdm import tqdm

parser = argparse.ArgumentParser(
    prog="docx_generator",
    description="A simple python script that allows batch generating docx files from template and keys and values stored in csv file.",
)

parser.add_argument(
    '-t',
    "--template",
    default="template.docx",
    help="Name of your template docx file, defaults to \"template.docx\". Keys to be replaced MUST be in this format: ${keyname}.",
    required=False,
    metavar="<template_name>"
)
parser.add_argument(
    "-c",
    "--csv",
    default="keys.csv",
    help="Name of your csv file, containing keys and values. Defaults to \"keys.csv\" First row should contain keys and should probably not include empty columns. Every row after the first one will be used to replace keys in template with provided values. If a key should contain a new line, mark it with {nl}.",
    required=False,
    metavar="<csv_name>"
)
parser.add_argument(
    "-s", 
    "--sep", 
    default=";", 
    help="Separator used for parsing csv file. Defaults to \";\"",
    required=False,
    metavar="<sep>"
)

args = parser.parse_args()

file = open(args.csv, "r", encoding="utf-8-sig")

line = file.readline()

keys = line.split(sep=args.sep)
for i in range(len(keys)):
    keys[i] = keys[i].strip()

dict_array = []
file_names = []

print("Parsing " + args.csv)

for line in file:
    d = {}
    tokens = line.split(";")
    for i in range(len(keys)):
        if(keys[i] == "output_file"):
            file_names.append(tokens[i])
            continue
        d[keys[i]] = tokens[i].strip().replace("{nl}", "\n")
    dict_array.append(d)

template = Document(args.template)

path =  pathlib.Path("./output")
path.mkdir(exist_ok=True)

print("Generating output")
for i in tqdm(range(len(dict_array))):
    doc = copy.deepcopy(template)
    docx_replace(doc, **dict_array[i])
    if(file_names[i] != ''):
        doc.save("./output/"+file_names[i]+".docx")
        continue
    doc.save("./output/"+str(i)+".docx")

print("Done!")