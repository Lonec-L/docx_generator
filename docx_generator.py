import copy
from docx import Document
from python_docx_replace import docx_replace

import argparse

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
print(args)
exit()

file = open(CSV_FILE_NAME, "r", encoding="utf-8-sig")

line = file.readline()

keys = line.split(sep=SEP)
for i in range(len(keys)):
    keys[i] = keys[i].strip()

dict_array = []


for line in file:
    d = {}
    tokens = line.split(";")
    for i in range(len(keys)):
        d[keys[i]] = tokens[i].strip().replace("{nl}", "\n")
    dict_array.append(d)

template = Document(TEMPLATE_NAME)

for i in range(len(dict_array)):
    doc = copy.deepcopy(template)
    docx_replace(doc, **dict_array[i])
    doc.save(str(i)+".docx")