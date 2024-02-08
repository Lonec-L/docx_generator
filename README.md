# docx_generator

A simple python script that allows batch generating docx files from template and csv file with keys & values.

## Installing and usage

First clone the repository in your desired location:
```
git clone https://github.com/Lonec-L/docx_generator.git
cd docx_generator
```

Create python venv:
```
python -m venv ./env
```
Activate the environment:

Windows: `.\env\Scripts\activate`

Linux: `./env/bin/activate`

<sup>If above doesnt work for you try searching for alternative for your system.</sup>

Install dependencies:
`pip install -r requirements.txt`

Run to display help:
`python docx_generator.py --help`

Use example:
`python docx_generator.py -t test_template.docx -c test_keys.csv -s ","`

## CSV specification

Key `output_file` is reserved for specifying names of output files. Otherwise the CSV file should look like this:

|key0|key1|key2|
|---|---|---|
|file1_key0_value|file1_key1_value|file1_key2_value|
|file2_key0_value|file2_key1_value|file2_key2_value|

Each key value from template will be replaced with corresponding value in output files (file[1-3]).

For example, for csv with the follwoing values:
|fruit|person|amount|
|---|---|---|
|banana|Tom|2|
|mango|Jack|3|
|kiwi|Larry|6|

We get 3 output documents, with keys `${fruit}`, `${person}` and `${amount}` replaced by values in rows. For Tom that would be:
```
${fruit} -> banana
${person} -> Tom
${amount} -> 2
```

## Output

Output files will be stored in folder `./output`. They are named with row number `[0-N]`, or name provided in csv file, column `output_file`.