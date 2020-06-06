The homepage of the scripts:  
[https://github.com/LeonidBeynenson/script_filling_rtfs_by_data_from_table]

# Filling RTF files by data from a CSV table

## General description

This repository contains the script `fill_rtf_templates.py` that allows filling RTF files using data from a CSV table.

To run the script place in this folder one .csv file in utf-8 encoding  
(typically the file is received from Google Sheets using Files -> Export -> Export as CSV)

Also the script searches in its folder one RTF file which name starts with underscore `_` -- it is the template.

The first line in the CSV file is the header -- it contains the names of the placeholders in the template that
will be substituted by values from the csv table.

Note that during the substitution the values from CSV file are converted from utf-8 encoding to RTF.

## Details on conversion utf-8 to RTF

The script `fill_rtf_templates.py` uses the files `alphabet_for_csv.txt` and `alphabet_for_rtf_ch2.rtf` to match
the corresponding symbols from utf-8 with the RTF encodings
(the file `alphabet_for_rtf_ch2.rtf` was edited manually in a binary files editor to have a form suitable for parsing).

Note that the script is intended for work with the Russian language, 
so I filled the alphabet files for cyrillic letters and several special chars.

Also note that the placeholderds should be written in the same language as the values in CSV files  
(it was a simple and rude approach to resolve the issue with RTF context)

## Usage

No parameters are required.
Use python3 to run the script `fill_rtf_templates.py`.

The script searches the RTF template and the CSV file in the folder where it is placed.

## Requirements

Checked on Windows Anaconda3(64-bit).

Also the script requires the package pprint, which may be installed on Anaconda by the command
```
conda install -c conda-forge pprintpp
```

# Converting RTF files to PDF files

## General description

Also this repository contains the script `convert_rtf_files_to_pdf.py` that allows converting all RTF in the
folder of the script to PDF files.

This script uses MS Word application for that.

For calling MS Word operations it uses pywin32 Python library to use the COM interface.

See documentation of win32gui library here  
[http://timgolden.me.uk/pywin32-docs/contents.html]

## Usage

No parameters are required.
Use python3 to run the script `convert_rtf_files_to_pdf.py`.

The script searches the RTF files in the folder where it is placed.

## Requirements

Checked on Windows Anaconda3(64-bit).

The pywin32 library may be installed by the command
```
conda install -c anaconda pywin32
```
