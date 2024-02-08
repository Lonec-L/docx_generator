# docx_generator

A simple python script that allows batch generating docx files from template and keys and values stored in csv file.

## Installing and usage

First clone the repository in your desired location:
```
git clone ...
cd docx_generator
```

Create python venv:
```
python -m venv ./env
```
Activate the environment:

Windows: `.\env\Scripts\activate`

Linux: `./env/bin/activate`
<sub>If above doesnt work for you try searching for alternative for your system.</sub>

Install dependencies:
`pip install -r requirements.txt`

Run to display help:
`python docx_generator.py --help`

Use example:
`python docx_generator.py -t test_template.docx -c test_keys.csv -s ","`